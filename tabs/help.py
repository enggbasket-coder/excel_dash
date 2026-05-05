"""
Help tab — moves the original left-sidebar "How to use" & "Important" notes
into a clean, documentation-style layout with FAQs, troubleshooting, and
keyboard shortcuts.
"""

from __future__ import annotations

import streamlit as st

from ui.theme import section_title, alert


# ───────────────────────────────────────────────────────────────
def _quick_start():
    section_title("Quick Start", hint="Get up and running in 4 steps")

    steps = [
        ("1", "📂", "Open your data file in Excel",
         "Open your <code>.xlsm</code> file (e.g. <code>SPS12.xlsx</code>) "
         "in Microsoft Excel. The dashboard cannot run macros until a workbook is open."),
        ("2", "🖱", "Come back here and click a macro button",
         "Switch to the <b>Macro</b> tab in this dashboard. Pick a quick preset "
         "or type any macro name in the Custom Macro box."),
        ("3", "⚙", "The macro runs inside your live workbook",
         "The macro executes in the background against your open workbook — "
         "no extra files are opened, no copies are made."),
        ("4", "📊", "Results appear in your workbook",
         "Once finished, results show up directly in your <code>.xlsm</code> file. "
         "The Macro tab also keeps a session log of every run."),
    ]

    cols = st.columns(2, gap="medium")
    for i, (n, icon, title, body) in enumerate(steps):
        with cols[i % 2]:
            st.markdown(
                f"""
                <div class="card card--accent" style="display:flex;gap:14px;align-items:flex-start;">
                    <div style="width:38px;height:38px;border-radius:12px;
                                background:linear-gradient(135deg,var(--pink-500),var(--pink-600));
                                color:#FFF;display:flex;align-items:center;justify-content:center;
                                font-weight:800;font-size:15px;flex-shrink:0;
                                box-shadow:var(--shadow-lg);">{n}</div>
                    <div>
                        <div style="font-weight:700;color:var(--ink-900);font-size:14px;">
                            {icon} &nbsp; {title}
                        </div>
                        <div style="color:var(--ink-500);font-size:13px;margin-top:6px;
                                    line-height:1.55;">{body}</div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# ───────────────────────────────────────────────────────────────
def _important_notes():
    section_title("Important", hint="Read before running for the first time")

    notes = [
        ("💻", "Microsoft Excel must be installed",
         "Excel must be installed on this machine, and your data file must be "
         "<b>open</b> in Excel — not just present on disk."),
        ("🍎", "macOS — first-run permission",
         "On the first run, macOS shows an <b>Automation</b> permission prompt "
         "asking the dashboard to control Excel. Click <b>OK</b>. If you miss "
         "it, enable it under <i>System Settings → Privacy &amp; Security → Automation</i>."),
        ("🪟", "Windows — enable macros",
         "Make sure macros are enabled in <i>Excel → File → Options → Trust Center → "
         "Trust Center Settings → Macro Settings</i>. The dashboard uses pywin32 "
         "to drive Excel via COM."),
        ("🔒", "Stays on your machine",
         "The dashboard only talks to your local Excel session. No data is uploaded, "
         "logged, or sent anywhere outside this computer."),
    ]

    for icon, title, body in notes:
        st.markdown(
            f"""
            <div class="card" style="display:flex;gap:14px;align-items:flex-start;">
                <div style="font-size:22px;line-height:1;">{icon}</div>
                <div>
                    <div style="font-weight:700;color:var(--ink-900);font-size:14px;">{title}</div>
                    <div style="color:var(--ink-500);font-size:13px;margin-top:4px;line-height:1.55;">
                        {body}
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ───────────────────────────────────────────────────────────────
def _faq():
    section_title("Frequently Asked Questions",
                  hint="Click a question to expand the answer")

    faqs = [
        ("Why does the dashboard say 'No Excel is open' even when Excel is running?",
         "The dashboard looks for an <b>active workbook</b>, not just the Excel app. "
         "Make sure you have at least one <code>.xlsm</code> or <code>.xlsx</code> file "
         "open and that you've clicked into its window once."),
        ("My macro errors out with 'Cannot run the macro'.",
         "Three usual causes:<br/>"
         "1. The macro name is wrong (it's case-sensitive on Windows-style references). "
         "Try using <code>Module1.MyMacro</code> instead of just <code>MyMacro</code>.<br/>"
         "2. Macros are disabled in Trust Center — re-enable them.<br/>"
         "3. The workbook is in protected view — click <b>Enable Editing</b> first."),
        ("Can I add more Quick Presets?",
         "Yes. Edit <code>tabs/macro.py</code> and append to the <code>QUICK_PRESETS</code> "
         "list at the top of the file. Each entry just needs a <code>name</code>, "
         "<code>icon</code>, and short <code>desc</code>."),
        ("Where are my Letters and Audit observations saved?",
         "Currently they live in the Streamlit session — they're cleared when you "
         "stop the app. To persist them across runs, hook the lists in "
         "<code>st.session_state</code> up to a SQLite/JSON store of your choice."),
        ("Can the same dashboard run on Linux?",
         "Not directly — the bridge to Excel uses macOS AppleScript or Windows COM. "
         "On Linux you'd need to run the macros via <code>libreoffice --headless</code> "
         "or remote-execute against an Excel host."),
    ]
    for q, a in faqs:
        with st.expander(q, expanded=False):
            st.markdown(a, unsafe_allow_html=True)


# ───────────────────────────────────────────────────────────────
def _troubleshooting():
    section_title("Troubleshooting", hint="Common errors and how to clear them")

    rows = [
        ("🔴 No Excel is open",
         "Open your <code>.xlsm</code> file in Excel and click <b>Refresh status</b> "
         "in the Macro tab."),
        ("🟡 osascript timed out",
         "Excel is busy or showing a modal dialog. Bring Excel to the front, "
         "dismiss any open dialogs, and retry."),
        ("🔴 pywin32 not installed (Windows)",
         "Run <code>pip install pywin32</code>, then restart the dashboard."),
        ("🟡 Permission denied (macOS)",
         "Open <i>System Settings → Privacy &amp; Security → Automation</i>, find "
         "this app (Terminal / iTerm / your launcher) and tick <b>Microsoft Excel</b>."),
    ]
    for title, body in rows:
        st.markdown(
            f"""
            <div class="card" style="padding:14px 18px;">
                <div style="font-weight:700;color:var(--ink-900);">{title}</div>
                <div style="color:var(--ink-500);font-size:13px;margin-top:4px;
                            line-height:1.55;">{body}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ───────────────────────────────────────────────────────────────
def _shortcuts_and_about():
    c1, c2 = st.columns(2, gap="medium")

    with c1:
        section_title("Keyboard Shortcuts")
        rows = [
            ("R", "Re-run the Streamlit app"),
            ("C", "Clear the on-screen cache"),
            ("Esc", "Stop the running script"),
            ("⌘ /  Ctrl + Enter", "Submit a focused text input"),
        ]
        kbd_html = "".join(
            f"""
            <div style="display:flex;justify-content:space-between;align-items:center;
                        padding:10px 0;border-bottom:1px solid var(--ink-100);">
                <div style="color:var(--ink-700);font-size:13.5px;">{label}</div>
                <kbd style="background:var(--ink-50);border:1px solid var(--ink-100);
                            color:var(--ink-700);padding:3px 9px;border-radius:8px;
                            font-family:ui-monospace,Menlo,monospace;font-size:12px;
                            font-weight:600;">{key}</kbd>
            </div>
            """ for key, label in rows
        )
        st.markdown(f'<div class="card">{kbd_html}</div>', unsafe_allow_html=True)

    with c2:
        section_title("About this dashboard")
        st.markdown(
            """
            <div class="card">
                <div style="font-weight:700;color:var(--ink-900);font-size:15px;">
                    Excel VBA Dashboard <span style="color:var(--pink-600);">v2.0</span>
                </div>
                <div style="color:var(--ink-500);font-size:13px;margin-top:6px;line-height:1.6;">
                    A modern, professional control surface for running Excel macros,
                    drafting EPC/contract correspondence, and conducting on-site audits —
                    without leaving the browser.<br/><br/>
                    Built on <b>Streamlit</b>. Talks to Excel via <b>AppleScript</b>
                    (macOS) or <b>pywin32</b> (Windows).
                </div>
                <div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap;">
                    <span class="meta-pill">Streamlit · 1.36+</span>
                    <span class="meta-pill">Python 3.10+</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ───────────────────────────────────────────────────────────────
def render_help_tab() -> None:
    alert(
        "info",
        "Welcome to the <b>Help Center</b>. Everything you need to get the most out of "
        "the Excel VBA Dashboard is below — quick start, important notes, FAQs, "
        "and troubleshooting.",
        icon="📚",
    )

    _quick_start()
    st.write("")
    _important_notes()
    st.write("")
    _faq()
    st.write("")
    _troubleshooting()
    st.write("")
    _shortcuts_and_about()
