"""
LLM Audit tab — AI-powered audit on any open Excel workbook.

Actions available:
  • RFI Duplicate Audit      — find duplicate RFI numbers across the whole workbook
  • Payment Procedure Audit  — compare payment procedure vs. quantity bill
  • Custom Prompt Audit      — user writes their own audit instruction

All LLM calls go to excel.epcmode.in (VPS proxy) — Google key never on this machine.
xlwings reads/writes any open Excel file without needing VBA inside it.
"""

from __future__ import annotations

from datetime import datetime

import streamlit as st

from core.vps_client  import (load_license_key, save_license_key,
                               validate_license, ask_llm, clear_license_key)
from core.excel_reader import (get_open_workbooks, get_sheets,
                                read_sheet_as_text, write_audit_column,
                                xlwings_available, excel_running,
                                xlwings_can_read, get_audit_column_options)
from ui.theme import section_title, alert

# ─────────────────────────────────────────────────────────────
#  Prompts
# ─────────────────────────────────────────────────────────────
_RFI_SYS = """You are an audit assistant for civil engineering projects.
You will receive a list of sheets from an Excel workbook, each containing RFI numbers.
Find ALL duplicate RFI numbers — across sheets AND within sheets.
For EVERY data row output exactly one line:
  SHEET <sheet_name> ROW <n>: <STATUS> - <short reason under 12 words>
STATUS = OK (unique) or DUPLICATE (seen before, name the earlier location).
No other text."""

_PAY_SYS = """You are an audit assistant for civil engineering billing.
You receive: (A) a Payment Procedure document and (B) Quantity of This Bill rows.
For EVERY bill row output exactly one line:
  ROW <n>: <STATUS> - <short reason under 15 words>
STATUS = OK or ISSUE.
No other text."""

_CUSTOM_SYS = "You are an expert Excel data auditor for civil engineering projects."

_QTY_SYS = """You are a quantity reconciliation auditor for civil engineering bills.

You receive ONE bill sheet. Each ROW <n> is a bill line item with columns that
typically include:
  - Scope, Amount, Unit rate
  - Quantity columns: "Upto Date", "Up to Previous", "This Bill", "Balance"
  - Workdone Value columns: "Upto Date", "Up to Previous", "This Bill"

For EVERY ROW <n>, verify the math:

  Check 1.  Quantity[Upto Date] = Quantity[Up to Previous] + Quantity[This Bill]
            (allow ±0.5% rounding)

  Check 2.  WorkdoneValue[This Bill] ≈ Quantity[This Bill] × Unit rate
            (allow ±1% rounding)

  Check 3.  WorkdoneValue[Upto Date] = WorkdoneValue[Up to Previous]
                                       + WorkdoneValue[This Bill]
            (allow ±1% rounding)

  Check 4.  If a row is a header / sub-total / blank, treat as OK (not billable).

Output STRICTLY one line per row, in order, nothing else:

  ROW <n>: <STATUS> - <reason under 18 words>

STATUS = OK  (all checks pass, OR header/empty row)
       = ISSUE  (any check fails — say which numbers don't add up)

Examples:
ROW 1: OK - Header row, not a billable item
ROW 2: OK - 0.088 × 35789090.91 = 3149440 matches workdone value
ROW 3: ISSUE - This Bill 0.238 but Upto - Previous = 0.244
ROW 4: ISSUE - Workdone 23798834 ≠ 0.238 × 99995104 = 23798835 (tolerance OK actually no this is fine)
ROW 5: OK - No claim this bill, balance preserved

Do NOT include any commentary, headings, or blank lines."""


# ─────────────────────────────────────────────────────────────
def _ensure_state():
    defaults = {
        "llm_license_key":     load_license_key(),
        "llm_license_info":    None,
        "llm_last_result":     None,
        "llm_workbook":        None,
        "llm_open_books":      [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


# ─────────────────────────────────────────────────────────────
#  License section
# ─────────────────────────────────────────────────────────────
def _license_section():
    section_title("License Key", hint="Your key is saved locally — enter once")

    key = st.session_state.llm_license_key
    info = st.session_state.llm_license_info

    if info and info.get("valid"):
        used  = info.get("calls_this_month", 0)
        quota = info.get("quota_monthly", "?")
        pct   = int(used / quota * 100) if isinstance(quota, int) and quota > 0 else 0
        expires = (' · Expires ' + info['expires_at']) if info.get('expires_at') else ''

        col1, col2 = st.columns([5, 1])
        with col1:
            st.markdown(
                f"""
                <div class="card" style="display:flex;align-items:center;gap:16px;padding:14px 18px;">
                    <div style="font-size:24px;">🟢</div>
                    <div>
                        <div style="font-weight:700;color:var(--ink-900);">
                            {info.get('customer_name','—')} &nbsp;
                            <span style="background:var(--pink-100);color:var(--pink-500);
                                  padding:2px 10px;border-radius:999px;font-size:12px;
                                  font-weight:600;">{info.get('plan','—').upper()}</span>
                        </div>
                        <div style="color:var(--ink-500);font-size:12.5px;margin-top:4px;">
                            {used} / {quota} calls this month ({pct}% used){expires}
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        with col2:
            st.write("")
            if st.button("🔑 Change", use_container_width=True, key="llm_change_key"):
                st.session_state.llm_license_info = None
                st.session_state.llm_license_key  = ""
                clear_license_key()
                st.rerun()
    else:
        c1, c2 = st.columns([4, 1])
        with c1:
            new_key = st.text_input(
                "License key",
                value=st.session_state.llm_license_key,
                placeholder="XXXX-XXXX-XXXX-XXXX",
                label_visibility="collapsed",
                key="llm_key_input",
            )
        with c2:
            if st.button("✅ Activate", type="primary", use_container_width=True,
                         key="llm_activate_btn"):
                if new_key.strip():
                    with st.spinner("Validating…"):
                        info = validate_license(new_key.strip())
                    if info.get("valid"):
                        save_license_key(new_key.strip())
                        st.session_state.llm_license_key  = new_key.strip()
                        st.session_state.llm_license_info = info
                        st.rerun()
                    else:
                        alert("err", info.get("message", "Invalid key."), icon="❌")
                else:
                    alert("warn", "Enter a license key.", icon="⚠️")


# ─────────────────────────────────────────────────────────────
#  Workbook picker
# ─────────────────────────────────────────────────────────────
def _workbook_picker() -> tuple[str, list[str]]:
    """Returns (selected_workbook_name, list_of_sheet_names)."""

    section_title("Select Workbook", hint="Choose from all open Excel files")

    # Always re-read on every render — workbooks change frequently
    books = get_open_workbooks()
    st.session_state.llm_open_books = books

    c2_col1, c2_col2 = st.columns([5, 1])
    with c2_col2:
        st.write("")
        if st.button("🔄 Refresh", use_container_width=True, key="refresh_books"):
            st.rerun()

    if not excel_running():
        alert("warn",
              "Excel is not running. Open a workbook in Excel, then click Refresh.",
              icon="⚠️")
        return "", []

    if not books:
        alert("warn",
              "No open workbooks found. Open a file in Excel then click Refresh.",
              icon="⚠️")
        return "", []

    # Workbook detected via AppleScript ✅ — but reading data needs xlwings,
    # so warn if xlwings can't connect (most common: Automation perm not granted)
    if not xlwings_available():
        alert("err",
              "xlwings not installed in this venv. "
              "Run: <code>.venv/bin/pip install xlwings</code> then restart Streamlit.",
              icon="❌")
        return "", []

    can_read, hint = xlwings_can_read()
    if not can_read:
        alert("warn", hint, icon="🔐")
        # Still let user select, in case xlwings recovers — but warn

    with c2_col1:
        wb = st.selectbox("Open workbooks", options=books,
                           index=0, key="llm_workbook_sel",
                           label_visibility="collapsed")
    sheets = get_sheets(wb) if wb else []
    return wb, sheets


# ─────────────────────────────────────────────────────────────
#  RFI duplicate audit
# ─────────────────────────────────────────────────────────────
def _rfi_action(wb: str, sheets: list[str]):
    section_title("RFI Duplicate Audit",
                  hint="Scans all sheets for duplicate RFI numbers — across sheets too")
    st.markdown(
        """<div class="card" style="color:var(--ink-500);font-size:13px;">
        Reads every sheet in the selected workbook, finds the RFI column (any header
        containing "RFI"), then uses AI to flag duplicates <b>within</b> and
        <b>across</b> sheets. Results written back to each sheet.
        </div>""", unsafe_allow_html=True)
    st.write("")

    if st.button("🚀 Run RFI Audit", type="primary", use_container_width=False):
        _run_rfi_audit(wb, sheets)


def _run_rfi_audit(wb: str, sheets: list[str]):
    if not st.session_state.llm_license_key:
        alert("err", "Enter your license key above first.", icon="❌")
        return

    # Build combined text from all sheets
    combined = ""
    sheet_maps: dict[str, tuple[list[int], int]] = {}

    with st.spinner("Reading workbook…"):
        for sh in sheets:
            text, row_map, hdr_row = read_sheet_as_text(wb, sh, char_budget=15000, max_rows=300)
            if text.strip() and "HEADER:" in text:
                # Only include sheets that look like they have RFI data
                if "rfi" in text.lower()[:200]:
                    combined += f"\n--- SHEET: {sh} ---\n{text}\n"
                    sheet_maps[sh] = (row_map, hdr_row)

    if not combined.strip():
        alert("warn", "No sheets with RFI columns found in this workbook.", icon="⚠️")
        return

    with st.spinner("Sending to AI (this takes ~10–30 seconds)…"):
        result = ask_llm(
            system_prompt=_RFI_SYS,
            user_prompt=f"Here is the workbook data:\n{combined}\nAudit for duplicates.",
            action="rfi_audit",
            license_key=st.session_state.llm_license_key,
        )

    if not result["ok"]:
        alert("err", f"AI error: {result['error']}", icon="❌")
        return

    st.session_state.llm_last_result = result["text"]
    _show_result_preview(result)

    # Parse and write back
    _write_rfi_results_to_excel(wb, sheet_maps, result["text"])


def _write_rfi_results_to_excel(wb, sheet_maps, reply):
    """Parse 'SHEET X ROW N: STATUS - reason' lines and write to each sheet."""
    lines = reply.strip().splitlines()
    from collections import defaultdict
    per_sheet: dict[str, dict[int, tuple[str, bool]]] = defaultdict(dict)

    for line in lines:
        line = line.strip()
        if not line.upper().startswith("SHEET"):
            continue
        try:
            # "SHEET <name> ROW <n>: <STATUS> - <reason>"
            rest = line[6:].strip()
            row_idx = rest.upper().find(" ROW ")
            sh_name = rest[:row_idx].strip()
            rest2   = rest[row_idx + 5:].strip()
            colon   = rest2.find(":")
            row_num = int(rest2[:colon].strip())
            verdict = rest2[colon + 1:].strip()
            is_ok   = verdict.upper().startswith("OK")
            per_sheet[sh_name][row_num] = (verdict, is_ok)
        except Exception:
            continue

    written = 0
    for sh_name, row_verdicts in per_sheet.items():
        if sh_name not in sheet_maps:
            continue
        row_map, hdr_row = sheet_maps[sh_name]
        results = []
        for i, _ in enumerate(row_map, start=1):
            if i in row_verdicts:
                results.append(row_verdicts[i])
            else:
                results.append(("Unique", True))

        ts = datetime.now().strftime("%Y-%m-%d %H:%M")
        r = write_audit_column(wb, sh_name, hdr_row, row_map, results,
                                column_title=f"RFI Audit ({ts})")
        if r["ok"]:
            written += 1

    if written:
        alert("ok", f"✅ Audit written to <b>{written}</b> sheet(s) in Excel.", icon="")
    else:
        alert("warn", "Could not write back to Excel. Is the file still open?", icon="⚠️")


# ─────────────────────────────────────────────────────────────
#  Payment procedure audit
# ─────────────────────────────────────────────────────────────
def _pay_action(wb: str, sheets: list[str]):
    section_title("Payment Procedure Audit",
                  hint="Compare payment procedure rules vs. billed quantities")

    c1, c2 = st.columns(2)
    with c1:
        proc_sheet = st.selectbox("Payment Procedure sheet", options=sheets,
                                   key="pay_proc_sheet")
    with c2:
        bill_opts = [s for s in sheets if s != proc_sheet] if proc_sheet else sheets
        bill_sheet = st.selectbox("Quantity of This Bill sheet", options=bill_opts,
                                   key="pay_bill_sheet")

    if proc_sheet and bill_sheet and proc_sheet == bill_sheet:
        alert("warn", "Select two different sheets.", icon="⚠️")
        return

    # ── Target column picker — only audit rows where THIS column has a value ──
    target_col = None
    if bill_sheet:
        col_opts, _ = get_audit_column_options(wb, bill_sheet)
        if col_opts:
            st.markdown(
                "<div style='color:var(--ink-500);font-size:13px;margin-top:6px;'>"
                "<b>Target column</b> — pick the column whose rows you want to audit. "
                "Only rows with a value in this column will be checked."
                "</div>", unsafe_allow_html=True,
            )
            labels = [o["label"] for o in col_opts]
            chosen = st.selectbox("Target column", options=labels,
                                   key="pay_target_col", label_visibility="collapsed")
            target_col = next((o["col_idx"] for o in col_opts if o["label"] == chosen), None)

    if st.button("🚀 Run Payment Audit", type="primary"):
        _run_pay_audit(wb, proc_sheet, bill_sheet, target_col)


def _run_pay_audit(wb, proc_sheet, bill_sheet, target_col=None):
    if not st.session_state.llm_license_key:
        alert("err", "Enter your license key above first.", icon="❌")
        return

    with st.spinner("Reading sheets…"):
        proc_text, _, _     = read_sheet_as_text(wb, proc_sheet,  char_budget=18000)
        bill_text, row_map, hdr_row = read_sheet_as_text(
            wb, bill_sheet, char_budget=30000, max_rows=300,
            target_col=target_col,
        )

    if not proc_text.strip():
        alert("warn", f"Sheet '{proc_sheet}' appears empty.", icon="⚠️")
        return
    if not bill_text.strip() or not row_map:
        alert("warn",
              f"No billable rows in '{bill_sheet}' for the selected target column. "
              "Pick a different column or check the data.",
              icon="⚠️")
        return

    user_prompt = (
        f"PAYMENT PROCEDURE (sheet '{proc_sheet}'):\n<<<\n{proc_text}\n>>>\n\n"
        f"QUANTITY OF THIS BILL (sheet '{bill_sheet}'):\n<<<\n{bill_text}\n>>>\n\n"
        "Now produce one  ROW <n>: <STATUS> - <reason>  line per bill row."
    )

    with st.spinner("Sending to AI…"):
        result = ask_llm(
            system_prompt=_PAY_SYS,
            user_prompt=user_prompt,
            action="pay_procedure",
            license_key=st.session_state.llm_license_key,
        )

    if not result["ok"]:
        alert("err", f"AI error: {result['error']}", icon="❌")
        return

    st.session_state.llm_last_result = result["text"]
    _show_result_preview(result)

    # Parse "ROW N: STATUS - reason" and write back
    results = _parse_row_verdicts(result["text"], len(row_map))
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    r  = write_audit_column(wb, bill_sheet, hdr_row, row_map, results,
                             column_title=f"Pay Audit ({ts})")
    if r["ok"]:
        alert("ok", f"✅ Audit written to Excel sheet <b>{bill_sheet}</b>.", icon="")
    else:
        alert("warn", f"Could not write to Excel: {r['error']}", icon="⚠️")


# ─────────────────────────────────────────────────────────────
#  Bill quantity audit (math reconciliation)
# ─────────────────────────────────────────────────────────────
def _qty_action(wb: str, sheets: list[str]):
    section_title("Bill Quantity Audit",
                  hint="Verify Quantity × Unit Rate = Workdone Value, and "
                       "Upto Date = Previous + This Bill")
    st.markdown(
        """<div class="card" style="color:var(--ink-500);font-size:13px;">
        For each billable row, the AI checks: <b>(1)</b> This Bill quantity adds
        correctly to previous + upto-date totals, <b>(2)</b> Quantity × Unit Rate
        equals Workdone Value, and <b>(3)</b> Workdone Value totals reconcile.
        Only rows that have a value in your selected target column are audited.
        </div>""", unsafe_allow_html=True)
    st.write("")

    bill_sheet = st.selectbox("Bill sheet", options=sheets, key="qty_bill_sheet")

    target_col = None
    if bill_sheet:
        col_opts, _ = get_audit_column_options(wb, bill_sheet)
        if col_opts:
            st.markdown(
                "<div style='color:var(--ink-500);font-size:13px;margin-top:6px;'>"
                "<b>Target column</b> — pick the column whose values to audit "
                "(typically <i>Quantity → This Bill</i>). Rows without a value "
                "in this column will be skipped."
                "</div>", unsafe_allow_html=True,
            )
            labels = [o["label"] for o in col_opts]
            chosen = st.selectbox("Target column", options=labels,
                                   key="qty_target_col", label_visibility="collapsed")
            target_col = next((o["col_idx"] for o in col_opts if o["label"] == chosen), None)

    if st.button("🚀 Run Quantity Audit", type="primary"):
        if not st.session_state.llm_license_key:
            alert("err", "Enter your license key above first.", icon="❌")
            return
        _run_qty_audit(wb, bill_sheet, target_col)


def _run_qty_audit(wb, bill_sheet, target_col=None):
    with st.spinner("Reading bill sheet…"):
        bill_text, row_map, hdr_row = read_sheet_as_text(
            wb, bill_sheet, char_budget=40000, max_rows=400,
            target_col=target_col,
        )
    if not bill_text.strip() or not row_map:
        alert("warn",
              f"No billable rows in '{bill_sheet}' for the selected target column.",
              icon="⚠️")
        return

    user_prompt = (
        f"BILL SHEET ('{bill_sheet}'):\n<<<\n{bill_text}\n>>>\n\n"
        "Now produce one  ROW <n>: <STATUS> - <reason>  line for EVERY row above. "
        "Do not skip any row. Output ONLY ROW lines."
    )

    with st.spinner("Sending to AI for math reconciliation…"):
        result = ask_llm(
            system_prompt=_QTY_SYS,
            user_prompt=user_prompt,
            action="qty_audit",
            license_key=st.session_state.llm_license_key,
        )

    if not result["ok"]:
        alert("err", f"AI error: {result['error']}", icon="❌")
        return

    st.session_state.llm_last_result = result["text"]
    _show_result_preview(result)

    results = _parse_row_verdicts(result["text"], len(row_map))
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    r  = write_audit_column(wb, bill_sheet, hdr_row, row_map, results,
                             column_title=f"Qty Audit ({ts})")
    if r["ok"]:
        alert("ok", f"✅ Quantity audit written to <b>{bill_sheet}</b>.", icon="")
    else:
        alert("warn", f"Could not write to Excel: {r['error']}", icon="⚠️")


# ─────────────────────────────────────────────────────────────
#  Custom prompt audit
# ─────────────────────────────────────────────────────────────
def _custom_action(wb: str, sheets: list[str]):
    section_title("Custom Prompt Audit", hint="Write your own audit instruction")

    audit_sheet = st.selectbox("Sheet to audit", options=sheets, key="custom_sheet")

    target_col = None
    if audit_sheet:
        col_opts, _ = get_audit_column_options(wb, audit_sheet)
        if col_opts:
            st.markdown(
                "<div style='color:var(--ink-500);font-size:13px;margin-top:6px;'>"
                "<b>Target column (optional)</b> — limit the audit to rows that "
                "have a value in this column. Pick the first option to audit all rows."
                "</div>", unsafe_allow_html=True,
            )
            labels = ["(audit all rows)"] + [o["label"] for o in col_opts]
            chosen = st.selectbox("Target column", options=labels,
                                   key="custom_target_col", label_visibility="collapsed")
            if chosen != "(audit all rows)":
                target_col = next((o["col_idx"] for o in col_opts if o["label"] == chosen), None)

    instruction = st.text_area(
        "Audit instruction",
        height=120,
        placeholder=(
            "E.g. 'Check if the quantity in column D matches the approved BOQ. "
            "For each row output: ROW N: OK - reason  or  ROW N: ISSUE - reason'"
        ),
        key="custom_instruction",
    )

    if st.button("🚀 Run Custom Audit", type="primary"):
        if not instruction.strip():
            alert("warn", "Enter an audit instruction.", icon="⚠️")
            return
        _run_custom_audit(wb, audit_sheet, instruction, target_col)


def _run_custom_audit(wb, sheet_name, instruction, target_col=None):
    if not st.session_state.llm_license_key:
        alert("err", "Enter your license key above first.", icon="❌")
        return

    with st.spinner("Reading sheet…"):
        text, row_map, hdr_row = read_sheet_as_text(
            wb, sheet_name, target_col=target_col,
        )

    if not text.strip() or not row_map:
        alert("warn",
              f"No rows to audit in '{sheet_name}' "
              f"{'for the selected target column' if target_col else ''}.",
              icon="⚠️")
        return

    user_prompt = (
        f"INSTRUCTION: {instruction}\n\n"
        f"SHEET DATA ('{sheet_name}'):\n<<<\n{text}\n>>>"
    )

    with st.spinner("Sending to AI…"):
        result = ask_llm(
            system_prompt=_CUSTOM_SYS,
            user_prompt=user_prompt,
            action="custom_audit",
            license_key=st.session_state.llm_license_key,
        )

    if not result["ok"]:
        alert("err", f"AI error: {result['error']}", icon="❌")
        return

    st.session_state.llm_last_result = result["text"]
    _show_result_preview(result)

    results = _parse_row_verdicts(result["text"], len(row_map))
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")
    r  = write_audit_column(wb, sheet_name, hdr_row, row_map, results,
                             column_title=f"AI Audit ({ts})")
    if r["ok"]:
        alert("ok", f"✅ Audit written to Excel sheet <b>{sheet_name}</b>.", icon="")
    else:
        alert("warn", f"Could not write to Excel: {r['error']}", icon="⚠️")


# ─────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────
def _parse_row_verdicts(reply: str, total_rows: int) -> list[tuple[str, bool]]:
    """Parse 'ROW N: STATUS - reason' into [(text, is_ok)] list."""
    verdicts: dict[int, tuple[str, bool]] = {}
    for line in reply.strip().splitlines():
        line = line.strip()
        upper = line.upper()
        if not upper.startswith("ROW "):
            continue
        try:
            colon = line.find(":")
            n     = int(line[4:colon].strip())
            rest  = line[colon + 1:].strip()
            dash  = rest.find("-")
            if dash >= 0:
                status = rest[:dash].strip().upper()
                reason = rest[dash + 1:].strip()
            else:
                status = rest.upper()
                reason = ""
            is_ok = status == "OK"
            text  = ("OK - " if is_ok else "Issue - ") + reason
            verdicts[n] = (text, is_ok)
        except Exception:
            continue

    return [
        verdicts.get(i + 1, ("No verdict returned", False))
        for i in range(total_rows)
    ]


def _show_result_preview(result: dict):
    """Show a compact summary + expander with raw AI reply."""
    remaining = result.get("calls_remaining", "?")
    lines     = result["text"].strip().splitlines()
    ok_count  = sum(1 for l in lines if ": OK" in l.upper())
    issue_count = sum(1 for l in lines if "ISSUE" in l.upper())

    st.markdown(
        f"""
        <div class="card" style="display:flex;gap:24px;padding:14px 20px;">
            <div style="text-align:center;">
                <div style="font-size:22px;">✅</div>
                <div style="font-weight:700;color:var(--good);">{ok_count}</div>
                <div style="color:var(--ink-500);font-size:12px;">OK</div>
            </div>
            <div style="text-align:center;">
                <div style="font-size:22px;">🔴</div>
                <div style="font-weight:700;color:var(--bad);">{issue_count}</div>
                <div style="color:var(--ink-500);font-size:12px;">Issues</div>
            </div>
            <div style="text-align:center;">
                <div style="font-size:22px;">🎟</div>
                <div style="font-weight:700;color:var(--ink-900);">{remaining}</div>
                <div style="color:var(--ink-500);font-size:12px;">Calls left</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.expander("📄 View raw AI response"):
        st.code(result["text"], language=None)


# ─────────────────────────────────────────────────────────────
#  Main render
# ─────────────────────────────────────────────────────────────
def render_llm_audit_tab() -> None:
    _ensure_state()

    # ── License ───────────────────────────────────────────────
    _license_section()

    # Validate key on first load if we have a saved key
    if (st.session_state.llm_license_key
            and st.session_state.llm_license_info is None):
        info = validate_license(st.session_state.llm_license_key)
        st.session_state.llm_license_info = info

    # Gate the rest of the tab on having a valid license
    if not (st.session_state.llm_license_info
            and st.session_state.llm_license_info.get("valid")):
        st.markdown(
            """<div class="card" style="text-align:center;color:var(--ink-400);padding:40px;">
            <div style="font-size:32px;margin-bottom:8px;">🔒</div>
            Enter a valid license key above to unlock AI audit features.
            </div>""",
            unsafe_allow_html=True,
        )
        return

    st.write("")

    # ── Workbook picker ───────────────────────────────────────
    wb, sheets = _workbook_picker()
    if not wb or not sheets:
        return

    st.write("")

    # ── Action selector ───────────────────────────────────────
    section_title("Select Audit Action", hint="Choose what you want the AI to check")

    ACTION_MAP = {
        "🔍  RFI Duplicate Audit":       "rfi",
        "💰  Payment Procedure Audit":   "pay",
        "🧮  Bill Quantity Audit":       "qty",
        "✏️   Custom Prompt Audit":       "custom",
    }
    action_label = st.radio(
        "Action",
        list(ACTION_MAP.keys()),
        horizontal=True,
        label_visibility="collapsed",
        key="llm_action_radio",
    )
    action = ACTION_MAP[action_label]

    st.write("")

    # ── Action-specific UI ────────────────────────────────────
    if action == "rfi":
        _rfi_action(wb, sheets)
    elif action == "pay":
        _pay_action(wb, sheets)
    elif action == "qty":
        _qty_action(wb, sheets)
    elif action == "custom":
        _custom_action(wb, sheets)
