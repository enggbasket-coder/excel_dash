"""
Macro tab — Excel status, quick presets (FindOverlap), custom macro runner,
and a recent-runs log. Replaces & professionalises the original UI.
"""

from __future__ import annotations

from datetime import datetime

import streamlit as st

from core.excel_bridge import check_status, run_macro
from ui.theme import section_title, alert


# Quick-preset macros — extend or load from settings later
QUICK_PRESETS: list[dict] = [
    {"name": "FindOverlapsAllSheets", "icon": "🔍", "desc": "Flag overlapping / duplicate chainages on every sheet"},
    {"name": "RFI_dupli",       "icon": "🧾", "desc": "Flag duplicate RFI numbers (same sheet & across sheets)"},
    {"name": "FormatSheet",     "icon": "🎨", "desc": "Apply standard SOR formatting"},
    {"name": "RecalcAll",       "icon": "🔄", "desc": "Force-recalculate every formula"},
    {"name": "ExportSummary",   "icon": "📤", "desc": "Dump summary to a new sheet"},
    {"name": "ValidateRanges",  "icon": "✅", "desc": "Check named ranges & references"},
    {"name": "BackupActive",    "icon": "💾", "desc": "Save a timestamped copy"},
]


# ───────────────────────────────────────────────────────────────
def _kpi_row(status: dict) -> None:
    """Top KPI strip — Excel state / Workbook / Macros run / Avg time."""
    log = st.session_state.get("macro_log", [])
    total = len(log)
    successful = sum(1 for r in log if r.get("ok"))

    col1, col2, col3, col4 = st.columns(4, gap="medium")

    with col1:
        state_label = "Connected" if status["is_open"] else ("Idle" if status["running"] else "Closed")
        state_icon = "🟢" if status["is_open"] else ("🟡" if status["running"] else "🔴")
        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi__icon">{state_icon}</div>
                <div class="kpi__label">Excel State</div>
                <div class="kpi__value">{state_label}</div>
                <div class="kpi__delta">{status.get('app') or 'No app detected'}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col2:
        wb = status.get("workbook") or "—"
        wb_short = wb if len(wb) <= 18 else wb[:16] + "…"
        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi__icon">📄</div>
                <div class="kpi__label">Active Workbook</div>
                <div class="kpi__value" style="font-size:18px;line-height:1.3">{wb_short}</div>
                <div class="kpi__delta">{'Live' if status['is_open'] else 'Open a file'}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col3:
        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi__icon">⚡</div>
                <div class="kpi__label">Macros Run</div>
                <div class="kpi__value">{total}</div>
                <div class="kpi__delta">{successful} successful</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col4:
        last = log[-1]["macro"] if log else "None yet"
        last_short = last if len(last) <= 16 else last[:14] + "…"
        st.markdown(
            f"""
            <div class="kpi">
                <div class="kpi__icon">🕘</div>
                <div class="kpi__label">Last Macro</div>
                <div class="kpi__value" style="font-size:18px;line-height:1.3">{last_short}</div>
                <div class="kpi__delta">{log[-1]['time'] if log else '—'}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ───────────────────────────────────────────────────────────────
def _status_card(status: dict) -> None:
    """Big status card (replaces the red banner from the original UI)."""
    if status["is_open"]:
        kind, icon = "ok", "✓"
        title = "Excel is connected"
        sub = f"Active workbook: <b>{status['workbook']}</b>"
    elif status["running"]:
        kind, icon = "warn", "!"
        title = "Excel is running but no workbook is open"
        sub = "Open your <b>.xlsm</b> file in Excel, then click <b>Refresh status</b>."
    else:
        kind, icon = "bad", "✕"
        title = "No Excel session detected"
        sub = "Open Excel and your <b>.xlsm</b> file, then click <b>Refresh status</b>."

    if status.get("error"):
        sub += f"<br/><span style='opacity:.75'>Error: {status['error']}</span>"

    st.markdown(
        f"""
        <div class="status-card status-card--{kind}">
            <div class="status-card__icon">{icon}</div>
            <div>
                <div class="status-card__title">{title}</div>
                <div class="status-card__sub">{sub}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ───────────────────────────────────────────────────────────────
def _record(macro: str, result: dict) -> None:
    st.session_state.macro_log.append({
        "macro": macro,
        "ok": result["ok"],
        "time": datetime.now().strftime("%H:%M:%S"),
        "date": datetime.now().strftime("%d %b"),
        "error": result.get("error") or "",
        "output": result.get("output") or "",
    })


# ───────────────────────────────────────────────────────────────
def _run_and_report(macro: str) -> None:
    with st.spinner(f"Running `{macro}` in Excel…"):
        result = run_macro(macro)
    _record(macro, result)
    if result["ok"]:
        alert("ok", f"Macro <b>{macro}</b> completed successfully.", icon="✅")
        if result.get("output"):
            with st.expander("Macro output"):
                st.code(result["output"])
    else:
        alert("err", f"Macro <b>{macro}</b> failed: {result['error']}", icon="❌")


# ───────────────────────────────────────────────────────────────
def render_macro_tab() -> None:
    # 1. Refresh status on every render so KPIs are accurate
    if "excel_status_cache" not in st.session_state:
        st.session_state.excel_status_cache = check_status()
    status = st.session_state.excel_status_cache

    # ── KPI strip ──────────────────────────────────────────────
    _kpi_row(status)
    st.write("")

    # ── Status card + refresh ──────────────────────────────────
    section_title("Excel Status", hint="Detect a running Excel session and active workbook")
    s_col, b_col = st.columns([5, 1])
    with s_col:
        _status_card(status)
    with b_col:
        st.write("")
        if st.button("🔄 Refresh status", use_container_width=True, type="primary"):
            st.session_state.excel_status_cache = check_status()
            st.rerun()

    st.write("")

    # ── Quick presets ──────────────────────────────────────────
    section_title("Quick Presets", hint="One-click macros for the most common jobs")

    rows = [QUICK_PRESETS[i:i + 3] for i in range(0, len(QUICK_PRESETS), 3)]
    for row in rows:
        cols = st.columns(3, gap="medium")
        for col, preset in zip(cols, row):
            with col:
                st.markdown(
                    f"""
                    <div class="macro-tile">
                        <div class="macro-tile__icon">{preset['icon']}</div>
                        <div class="macro-tile__name">{preset['name']}</div>
                        <div class="macro-tile__desc">{preset['desc']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button(f"▶ Run {preset['name']}", key=f"preset_{preset['name']}",
                             use_container_width=True):
                    if not status["is_open"]:
                        alert("warn",
                              "Open your <b>.xlsm</b> file in Excel first, then run the macro.",
                              icon="⚠️")
                    else:
                        _run_and_report(preset["name"])

    st.write("")

    # ── Custom macro ───────────────────────────────────────────
    section_title("Custom Macro", hint="Run any macro by name from the active workbook")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        c1, c2 = st.columns([4, 1])
        with c1:
            macro_name = st.text_input(
                "Macro name",
                key="custom_macro_name",
                placeholder="e.g. ProcessSheet, GenerateReport, Module1.MyMacro",
                label_visibility="collapsed",
            )
        with c2:
            run_clicked = st.button("🚀 Run Macro", use_container_width=True, type="primary")
        st.markdown('</div>', unsafe_allow_html=True)

        if run_clicked:
            if not status["is_open"]:
                alert("warn",
                      "Open your <b>.xlsm</b> file in Excel first, then run the macro.",
                      icon="⚠️")
            elif not macro_name.strip():
                alert("warn", "Enter a macro name to run.", icon="⚠️")
            else:
                _run_and_report(macro_name.strip())

    st.write("")

    # ── Recent runs log ────────────────────────────────────────
    section_title("Recent Runs",
                  hint=f"{len(st.session_state.macro_log)} entries in this session")
    log = st.session_state.macro_log
    if not log:
        st.markdown(
            """
            <div class="card" style="text-align:center;color:var(--ink-500);padding:32px;">
                <div style="font-size:28px;margin-bottom:6px;">📭</div>
                No macros have been run yet in this session.<br/>
                Try a Quick Preset or enter a custom macro name above.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        for entry in reversed(log[-10:]):
            badge = ('<span style="background:#DCFCE7;color:#166534;padding:3px 10px;'
                     'border-radius:999px;font-weight:600;font-size:12px;">SUCCESS</span>'
                     if entry["ok"] else
                     '<span style="background:#FEE2E2;color:#991B1B;padding:3px 10px;'
                     'border-radius:999px;font-weight:600;font-size:12px;">FAILED</span>')
            err_html = (f'<div style="margin-top:6px;color:#991B1B;font-size:12.5px;">'
                        f'<b>Error:</b> {entry["error"]}</div>') if entry["error"] else ""
            st.markdown(
                f"""
                <div class="card" style="margin-bottom:10px;">
                    <div style="display:flex;align-items:center;justify-content:space-between;">
                        <div>
                            <div style="font-weight:700;color:var(--ink-900);">{entry['macro']}</div>
                            <div style="color:var(--ink-500);font-size:12px;">
                                {entry['date']} · {entry['time']}
                            </div>
                        </div>
                        {badge}
                    </div>
                    {err_html}
                </div>
                """,
                unsafe_allow_html=True,
            )

        if st.button("🗑 Clear log", key="clear_macro_log"):
            st.session_state.macro_log = []
            st.rerun()
