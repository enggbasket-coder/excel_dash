"""
Excel VBA Dashboard — Redesigned
A modern, professional Streamlit dashboard for running Excel macros, drafting
letters, conducting audits, and accessing help — all in one polished UI.
"""

from __future__ import annotations

import platform
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import streamlit as st

from ui.theme import inject_theme, inject_floating_menu_button
from ui.sidebar import render_sidebar
from tabs.macro import render_macro_tab
from tabs.letters import render_letters_tab
from tabs.audit import render_audit_tab
from tabs.help import render_help_tab
from tabs.llm_audit import render_llm_audit_tab
from tabs.stripchart import render_stripchart_tab


# ────────────────────────────────────────────────────────────────────────────
# Page config — must be the very first Streamlit call
# ────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Excel VBA Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inject the custom Koki-inspired theme
inject_theme()

# Inject a top-left floating "Menu" button as a fallback for re-opening the
# sidebar after the user has collapsed it. Uses streamlit.components.v1.html
# so it actually executes (st.markdown sandboxes <script> tags). Single
# injection — no setInterval, no MutationObserver — so it cannot loop.
inject_floating_menu_button()


# ────────────────────────────────────────────────────────────────────────────
# Session state defaults
# ────────────────────────────────────────────────────────────────────────────
def _init_state() -> None:
    defaults = {
        "active_tab": "Macro",
        "excel_status": None,           # None | "open" | "closed"
        "excel_workbook": None,         # path/name of detected workbook
        "macro_log": [],                # list[dict] — last macro runs
        "letter_drafts": [],            # list[dict] — saved drafts
        "audit_observations": [],       # list[dict] — audit entries
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


_init_state()


# ────────────────────────────────────────────────────────────────────────────
# Sidebar (brand + nav)
# ────────────────────────────────────────────────────────────────────────────
TABS = ["Macro", "Letters", "Audit", "LLM Audit", "StripChart", "Help"]
active = render_sidebar(TABS, st.session_state.active_tab)
st.session_state.active_tab = active


# ────────────────────────────────────────────────────────────────────────────
# Top header bar
# ────────────────────────────────────────────────────────────────────────────
header_left, header_right = st.columns([6, 2])
with header_left:
    st.markdown(
        f"""
        <div class="page-header">
            <div class="page-header__greeting">Good {('morning' if datetime.now().hour < 12 else 'afternoon' if datetime.now().hour < 17 else 'evening')}, Raj</div>
            <div class="page-header__title">{active}</div>
            <div class="page-header__subtitle">Excel VBA Dashboard · Open your data file in Excel · click a button here · macro runs in your live workbook</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with header_right:
    st.markdown(
        f"""
        <div class="page-header__meta">
            <div class="meta-pill"><span class="meta-pill__dot"></span>{platform.system()} · Python {sys.version_info.major}.{sys.version_info.minor}</div>
            <div class="meta-date">{datetime.now().strftime('%A, %d %b %Y')}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="page-divider"></div>', unsafe_allow_html=True)


# ────────────────────────────────────────────────────────────────────────────
# Route to the active tab
# ────────────────────────────────────────────────────────────────────────────
if active == "Macro":
    render_macro_tab()
elif active == "Letters":
    render_letters_tab()
elif active == "Audit":
    render_audit_tab()
elif active == "LLM Audit":
    render_llm_audit_tab()
elif active == "StripChart":
    render_stripchart_tab()
elif active == "Help":
    render_help_tab()
