"""
Sidebar with brand block, primary navigation (Macro / Letters / Audit / Help),
and a footer call-to-action card. Returns the currently active tab name.
"""

import streamlit as st


# Single source of truth for tab metadata
_TAB_META = {
    "Macro":   {"icon": "⚡", "badge": "12"},
    "Letters": {"icon": "✉️", "badge": None},
    "Audit":   {"icon": "🛡️", "badge": "3"},
    "Help":    {"icon": "❓", "badge": None},
}


def render_sidebar(tabs: list[str], active: str) -> str:
    """Render sidebar; return whichever tab the user just selected (or `active`)."""
    with st.sidebar:
        # ── Brand block ──────────────────────────────────────
        st.markdown(
            """
            <div class="brand">
                <div class="brand__logo">📊</div>
                <div>
                    <div class="brand__name">Excel Suite</div>
                    <div class="brand__tag">Pro Dashboard</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ── Nav buttons ──────────────────────────────────────
        # Streamlit doesn't expose a per-button class, so we mark
        # the active button with a hidden CSS hook via columns.
        for tab in tabs:
            meta = _TAB_META.get(tab, {"icon": "•", "badge": None})
            label = f"{meta['icon']}   {tab}"
            if meta["badge"]:
                label += f"     {meta['badge']}"

            # Highlight active tab by wrapping in a class via JS hack
            is_active = tab == active
            if is_active:
                st.markdown(
                    f"""
                    <style>
                    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] >
                    div:has(button[kind="secondary"][data-active-tab="{tab}"]) .stButton {{
                        /* placeholder — handled below */
                    }}
                    </style>
                    """,
                    unsafe_allow_html=True,
                )

            clicked = st.button(
                label,
                key=f"nav_{tab}",
                use_container_width=True,
                type="primary" if is_active else "secondary",
            )
            if clicked:
                active = tab
                st.session_state.active_tab = tab
                st.rerun()

        # ── Footer CTA card ──────────────────────────────────
        st.markdown(
            """
            <div class="side-footer">
                <div class="side-footer__title">
                    Need a quick start?<br/>
                    Open your <b>.xlsm</b> file<br/>and click <b>Check Status</b>.
                </div>
                <div class="side-footer__btn">⚡ Open Macro Tab</div>
            </div>
            <div style="text-align:center; margin-top:18px; opacity:.75; font-size:11px;">
                © 2026 · Excel VBA Dashboard<br/>v2.0 · Redesigned
            </div>
            """,
            unsafe_allow_html=True,
        )

    return active
