"""
Theme for the Excel VBA Dashboard — mint canvas + forest-green accents.
Note: the CSS variables are still named `--pink-*` for backwards compatibility
with the original Koki-inspired layout; the hex values themselves are green.
Modern card system, soft shadows, rounded corners, professional typography.
"""

import streamlit as st


_CSS = """
<style>
/* ─────────────────────────────────────────────────────────
   Design tokens
   ───────────────────────────────────────────────────────── */
:root {
    /* Mint / forest-green palette (Chrome new-tab inspired) */
    --pink-50:  #E8F4EC;   /* mint canvas tint    */
    --pink-100: #D2EADB;   /* card border / chips */
    --pink-200: #A9D9BA;   /* dashed borders      */
    --pink-400: #2B8A6B;   /* hover accent        */
    --pink-500: #0E6A4E;   /* PRIMARY forest      */
    --pink-600: #0B5740;   /* primary hover       */
    --pink-700: #084632;   /* darkest accent      */

    --ink-900: #0F1F1A;
    --ink-700: #2C3E37;
    --ink-500: #5C6B65;
    --ink-300: #A8B5AE;
    --ink-100: #E2E9E5;
    --ink-50:  #F2F7F4;

    --good:    #15803D;
    --warn:    #B45309;
    --bad:     #B91C1C;

    --radius-sm: 10px;
    --radius:    16px;
    --radius-lg: 22px;

    --shadow-sm: 0 1px 2px rgba(15, 31, 26, 0.05);
    --shadow:    0 4px 14px rgba(15, 31, 26, 0.07);
    --shadow-lg: 0 10px 30px rgba(14, 106, 78, 0.14);
}

/* Import the same font Koki uses */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
    color: var(--ink-900);
}

/* ─────────────────────────────────────────────────────────
   Global canvas
   ───────────────────────────────────────────────────────── */
.stApp {
    background: linear-gradient(180deg, #E8F4EC 0%, #F4FAF6 320px, #FFFFFF 100%);
}

/* Hide Streamlit's default chrome we don't want */
#MainMenu, footer {
    visibility: hidden;
    height: 0;
}

/* Header bar — keep it (so Streamlit's native sidebar-collapse arrow remains
   visible & clickable), just transparent + hide the Deploy / kebab controls.
   IMPORTANT: do NOT add `visibility:hidden` or position/display overrides on
   the collapsed-sidebar control — Streamlit toggles it via React, and forcing
   visibility creates a fight-loop where clicking opens then immediately closes
   the sidebar. */
header[data-testid="stHeader"] {
    background: transparent !important;
    box-shadow: none !important;
    border: none !important;
}
header[data-testid="stHeader"] [data-testid="stToolbar"] {
    visibility: hidden;
}

/* Style ONLY the visual appearance of the open-sidebar close button (X / <)
   so it reads against the dark-green sidebar. No display or visibility rules. */
section[data-testid="stSidebar"] button[kind="header"],
section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] button {
    color: #FFFFFF !important;
    opacity: 0.90 !important;
}

/* Streamlit's collapsed-sidebar arrow — make it a dark-green pill so it's
   clearly visible against the mint background. ONLY visual styling here:
   no `display`, `visibility`, or `position` overrides, so we don't fight
   Streamlit's React state management. */
[data-testid="stSidebarCollapsedControl"] button,
[data-testid="collapsedControl"] button,
button[data-testid="stExpandSidebarButton"],
button[data-testid="baseButton-headerNoPadding"][aria-label*="sidebar" i] {
    background: linear-gradient(135deg, #0E6A4E, #084632) !important;
    color: #FFFFFF !important;
    border-radius: 12px !important;
    box-shadow: 0 6px 18px rgba(14, 106, 78, 0.30) !important;
    border: none !important;
    min-width: 38px !important;
    min-height: 38px !important;
}
[data-testid="stSidebarCollapsedControl"] button svg,
[data-testid="collapsedControl"] button svg,
button[data-testid="stExpandSidebarButton"] svg {
    color: #FFFFFF !important;
    fill: #FFFFFF !important;
    width: 18px !important;
    height: 18px !important;
}
[data-testid="stSidebarCollapsedControl"] button:hover,
[data-testid="collapsedControl"] button:hover,
button[data-testid="stExpandSidebarButton"]:hover {
    background: linear-gradient(135deg, #0B5740, #084632) !important;
    box-shadow: 0 10px 24px rgba(14, 106, 78, 0.45) !important;
}

.block-container {
    padding-top: 1.4rem !important;
    padding-bottom: 3rem !important;
    max-width: 1320px;
}

/* ─────────────────────────────────────────────────────────
   Sidebar — deep forest green
   ───────────────────────────────────────────────────────── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0E6A4E 0%, #084632 100%);
    border-right: none;
    box-shadow: 6px 0 24px rgba(8, 70, 50, 0.18);
}
section[data-testid="stSidebar"] > div {
    background: transparent;
}
section[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}

/* Brand block */
.brand {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 8px 4px 24px 4px;
    border-bottom: 1px solid rgba(255, 255, 255, 0.18);
    margin-bottom: 18px;
}
.brand__logo {
    width: 44px; height: 44px;
    border-radius: 12px;
    background: rgba(255, 255, 255, 0.18);
    display: flex; align-items: center; justify-content: center;
    font-size: 22px;
    backdrop-filter: blur(4px);
}
.brand__name {
    font-size: 20px;
    font-weight: 800;
    letter-spacing: -0.01em;
    line-height: 1;
}
.brand__tag {
    font-size: 11px;
    opacity: 0.80;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin-top: 2px;
}

/* Nav buttons (Streamlit buttons inside sidebar) */
section[data-testid="stSidebar"] .stButton > button {
    width: 100%;
    background: transparent;
    color: #FFFFFF !important;
    border: none;
    border-radius: 12px;
    padding: 12px 16px;
    text-align: left;
    font-weight: 500;
    font-size: 14.5px;
    transition: all 0.15s ease;
    margin-bottom: 4px;
    box-shadow: none;
    justify-content: flex-start;
    display: flex;
}
section[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255, 255, 255, 0.14);
    transform: translateX(2px);
    border: none;
}
section[data-testid="stSidebar"] .stButton > button:focus {
    box-shadow: none !important;
    outline: none !important;
}
/* Active nav state — primary buttons inside the sidebar render as a white pill.
   Streamlit wraps the label in a nested <p>/<span>/<div>, so we have to push
   the color down explicitly or the inner text stays white-on-white. */
section[data-testid="stSidebar"] .stButton > button[kind="primary"],
section[data-testid="stSidebar"] [data-testid="stBaseButton-primary"] {
    background: #FFFFFF !important;
    color: var(--pink-600) !important;
    font-weight: 700;
    border: none !important;
    box-shadow: 0 6px 16px rgba(0,0,0,0.12) !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] *,
section[data-testid="stSidebar"] [data-testid="stBaseButton-primary"] * {
    color: var(--pink-600) !important;
    fill: var(--pink-600) !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover,
section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover *,
section[data-testid="stSidebar"] [data-testid="stBaseButton-primary"]:hover,
section[data-testid="stSidebar"] [data-testid="stBaseButton-primary"]:hover * {
    background: #FFFFFF !important;
    color: var(--pink-700) !important;
    fill: var(--pink-700) !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"]:hover {
    transform: translateX(2px);
}

/* Sidebar footer card */
.side-footer {
    margin-top: 28px;
    background: rgba(255, 255, 255, 0.12);
    border: 1px dashed rgba(255, 255, 255, 0.30);
    border-radius: 14px;
    padding: 16px;
    text-align: center;
    backdrop-filter: blur(4px);
}
.side-footer__title { font-size: 12.5px; opacity: 0.92; line-height: 1.4; }
.side-footer__btn {
    margin-top: 10px;
    background: #FFFFFF;
    color: var(--pink-600);
    border-radius: 999px;
    padding: 8px 16px;
    font-weight: 700;
    font-size: 13px;
    display: inline-block;
}

/* ─────────────────────────────────────────────────────────
   Page header
   ───────────────────────────────────────────────────────── */
.page-header__greeting {
    color: var(--ink-500);
    font-size: 13px;
    font-weight: 500;
    margin-bottom: 4px;
}
.page-header__title {
    font-size: 28px;
    font-weight: 800;
    color: var(--ink-900);
    letter-spacing: -0.02em;
    line-height: 1.15;
}
.page-header__subtitle {
    color: var(--ink-500);
    font-size: 13.5px;
    margin-top: 6px;
    max-width: 720px;
}
.page-header__meta {
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 8px;
    padding-top: 6px;
}
.meta-pill {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: var(--pink-50);
    color: var(--pink-600);
    border: 1px solid var(--pink-100);
    border-radius: 999px;
    padding: 6px 14px;
    font-size: 12px;
    font-weight: 600;
}
.meta-pill__dot {
    width: 8px; height: 8px; border-radius: 50%;
    background: var(--pink-500);
    box-shadow: 0 0 0 3px rgba(14, 106, 78, 0.20);
}
.meta-date {
    font-size: 12px; color: var(--ink-500); font-weight: 500;
}
.page-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--ink-100), transparent);
    margin: 20px 0 24px 0;
}

/* ─────────────────────────────────────────────────────────
   Section titles
   ───────────────────────────────────────────────────────── */
.section-title {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: 6px 0 14px 0;
}
.section-title__left {
    display: flex;
    align-items: center;
    gap: 10px;
}
.section-title__bar {
    width: 4px; height: 22px;
    background: var(--pink-500);
    border-radius: 4px;
}
.section-title__text {
    font-size: 17px;
    font-weight: 700;
    color: var(--ink-900);
    letter-spacing: -0.01em;
}
.section-title__hint {
    color: var(--ink-500);
    font-size: 12px;
    font-weight: 500;
}

/* ─────────────────────────────────────────────────────────
   Cards
   ───────────────────────────────────────────────────────── */
.card {
    background: #FFFFFF;
    border: 1px solid var(--ink-100);
    border-radius: var(--radius);
    padding: 22px 24px;
    box-shadow: var(--shadow-sm);
    transition: box-shadow 0.15s ease, transform 0.15s ease;
    margin-bottom: 16px;
}
.card:hover {
    box-shadow: var(--shadow);
}
.card--accent {
    border-left: 4px solid var(--pink-500);
}
.card__title {
    font-size: 14px;
    font-weight: 700;
    color: var(--ink-900);
    margin: 0 0 4px 0;
    letter-spacing: -0.005em;
}
.card__sub {
    font-size: 12.5px;
    color: var(--ink-500);
    margin-bottom: 16px;
}

/* Status card (Excel open/closed) */
.status-card {
    display: flex;
    align-items: center;
    gap: 18px;
    background: #FFFFFF;
    border: 1px solid var(--ink-100);
    border-radius: var(--radius);
    padding: 20px 24px;
    box-shadow: var(--shadow-sm);
}
.status-card__icon {
    width: 52px; height: 52px;
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-size: 26px;
    flex-shrink: 0;
}
.status-card--ok .status-card__icon   { background: #DCFCE7; color: var(--good); }
.status-card--bad .status-card__icon  { background: #FEE2E2; color: var(--bad); }
.status-card--warn .status-card__icon { background: #FEF3C7; color: var(--warn); }
.status-card__title { font-weight: 700; font-size: 15px; color: var(--ink-900); }
.status-card__sub   { font-size: 13px; color: var(--ink-500); margin-top: 2px; }

/* KPI card */
.kpi {
    background: #FFFFFF;
    border: 1px solid var(--ink-100);
    border-radius: var(--radius);
    padding: 18px 20px;
    box-shadow: var(--shadow-sm);
    height: 100%;
}
.kpi__label {
    font-size: 12px;
    color: var(--ink-500);
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 600;
}
.kpi__value {
    font-size: 28px;
    font-weight: 800;
    color: var(--ink-900);
    letter-spacing: -0.02em;
    margin-top: 8px;
    line-height: 1;
}
.kpi__delta {
    font-size: 12px;
    margin-top: 8px;
    display: inline-flex;
    align-items: center;
    gap: 4px;
    background: var(--pink-50);
    color: var(--pink-600);
    border-radius: 999px;
    padding: 3px 10px;
    font-weight: 600;
}
.kpi__icon {
    float: right;
    width: 38px; height: 38px;
    border-radius: 12px;
    background: var(--pink-50);
    color: var(--pink-600);
    display: flex; align-items: center; justify-content: center;
    font-size: 18px;
    margin-top: -4px;
}

/* Macro preset tile */
.macro-tile {
    background: #FFFFFF;
    border: 1px solid var(--ink-100);
    border-radius: var(--radius);
    padding: 18px;
    text-align: center;
    box-shadow: var(--shadow-sm);
    transition: all 0.15s ease;
}
.macro-tile:hover {
    box-shadow: var(--shadow-lg);
    transform: translateY(-2px);
    border-color: var(--pink-200);
}
.macro-tile__icon {
    width: 46px; height: 46px;
    margin: 0 auto 10px;
    border-radius: 12px;
    background: linear-gradient(135deg, var(--pink-500), var(--pink-600));
    color: #FFF;
    display: flex; align-items: center; justify-content: center;
    font-size: 22px;
    box-shadow: var(--shadow-lg);
}
.macro-tile__name { font-weight: 700; font-size: 14px; color: var(--ink-900); }
.macro-tile__desc { font-size: 12px; color: var(--ink-500); margin-top: 2px; }

/* ─────────────────────────────────────────────────────────
   Buttons (main area)
   ───────────────────────────────────────────────────────── */
.stButton > button {
    border-radius: 12px;
    font-weight: 600;
    font-size: 14px;
    padding: 9px 20px;
    border: 1px solid var(--ink-100);
    background: #FFFFFF;
    color: var(--ink-700);
    transition: all 0.15s ease;
    box-shadow: var(--shadow-sm);
}
.stButton > button:hover {
    border-color: var(--pink-400);
    color: var(--pink-600);
    box-shadow: var(--shadow);
    transform: translateY(-1px);
}
.stButton > button:focus {
    outline: none !important;
    box-shadow: 0 0 0 3px rgba(14, 106, 78, 0.20) !important;
    color: var(--pink-600);
    border-color: var(--pink-400);
}
/* Primary button (kind="primary") */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, var(--pink-500), var(--pink-600));
    color: #FFFFFF !important;
    border: none;
    box-shadow: 0 6px 18px rgba(14, 106, 78, 0.30);
}
.stButton > button[kind="primary"]:hover {
    transform: translateY(-1px);
    box-shadow: 0 10px 24px rgba(14, 106, 78, 0.40);
    color: #FFFFFF !important;
}

/* Inputs */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea,
.stSelectbox > div > div,
.stDateInput > div > div > input {
    border-radius: 12px !important;
    border: 1px solid var(--ink-100) !important;
    background: #FFFFFF !important;
    font-size: 14px !important;
}
.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: var(--pink-400) !important;
    box-shadow: 0 0 0 3px rgba(14, 106, 78, 0.12) !important;
}

/* File uploader */
[data-testid="stFileUploader"] section {
    background: var(--pink-50);
    border: 2px dashed var(--pink-200);
    border-radius: var(--radius);
    padding: 18px;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--pink-400);
}

/* Metrics — Streamlit native st.metric */
[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid var(--ink-100);
    border-radius: var(--radius);
    padding: 16px 18px;
    box-shadow: var(--shadow-sm);
}
[data-testid="stMetricLabel"] {
    color: var(--ink-500);
    font-weight: 600;
    text-transform: uppercase;
    font-size: 11px;
    letter-spacing: 0.08em;
}
[data-testid="stMetricValue"] {
    color: var(--ink-900);
    font-weight: 800;
    font-size: 28px;
}

/* Expanders */
.streamlit-expanderHeader {
    border-radius: 12px !important;
    background: #FFFFFF !important;
    border: 1px solid var(--ink-100) !important;
    font-weight: 600 !important;
}

/* Tables */
.stDataFrame {
    border-radius: var(--radius);
    overflow: hidden;
    border: 1px solid var(--ink-100);
}

/* Code blocks (for macro logs) */
pre, code {
    background: #0F1F1A !important;
    color: #A9D9BA !important;
    border-radius: 10px !important;
    font-size: 12.5px !important;
}

/* Alert boxes */
.alert {
    padding: 14px 18px;
    border-radius: var(--radius);
    display: flex;
    align-items: flex-start;
    gap: 12px;
    margin: 8px 0 16px 0;
    font-size: 13.5px;
    font-weight: 500;
    border: 1px solid;
}
.alert--info  { background: #EFF6FF; border-color: #DBEAFE; color: #1D4ED8; }
.alert--ok    { background: #F0FDF4; border-color: #BBF7D0; color: #166534; }
.alert--warn  { background: #FFFBEB; border-color: #FDE68A; color: #92400E; }
.alert--err   { background: #FEF2F2; border-color: #FECACA; color: #991B1B; }
.alert__icon  { font-size: 18px; line-height: 1; }
.alert__body  { line-height: 1.5; }
.alert__body strong { font-weight: 700; }
</style>
"""


def inject_theme() -> None:
    """Inject the global stylesheet. Call once after st.set_page_config."""
    st.markdown(_CSS, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
# Floating "open menu" button
# ─────────────────────────────────────────────────────────────────
# Why this is built with streamlit.components.v1.html and NOT st.markdown:
#   - st.markdown(unsafe_allow_html=True) sandboxes / strips <script> tags
#   - components.v1.html renders inside an <iframe> where <script> DOES run
#   - We escape the iframe with window.parent.document to find Streamlit's
#     real sidebar control on the host page
#   - We bind ONE click handler — no setInterval / MutationObserver, so the
#     toggle-loop bug from the previous attempt cannot recur.
_FLOATING_MENU_HTML = """
<script>
  (function () {
      // We are inside Streamlit's component iframe.
      // Inject the floating button into the PARENT document so it's visible
      // (a `position:fixed` element inside a 0-height iframe gets clipped).
      const hostDoc = window.parent && window.parent.document;
      if (!hostDoc) return;

      // Reuse the existing button if it's already on the host page so we
      // don't duplicate it on every Streamlit rerun.
      let btn = hostDoc.getElementById('exsuite-floating-menu');
      if (!btn) {
          btn = hostDoc.createElement('button');
          btn.id = 'exsuite-floating-menu';
          btn.title = 'Toggle sidebar';
          btn.innerHTML = '☰';
          btn.style.cssText = [
              'position: fixed',
              'top: 12px',
              'left: 12px',
              'z-index: 999999',
              'width: 36px',
              'height: 36px',
              'border-radius: 10px',
              'border: none',
              'background: linear-gradient(135deg, #0E6A4E, #084632)',
              'color: #FFFFFF',
              'font-size: 17px',
              'font-weight: 700',
              'cursor: pointer',
              'box-shadow: 0 4px 12px rgba(14, 106, 78, 0.35)',
              'display: flex',
              'align-items: center',
              'justify-content: center',
              'opacity: 0.95',
              'transition: transform .15s ease, box-shadow .15s ease, opacity .15s ease',
          ].join(';');
          btn.addEventListener('mouseover', () => {
              btn.style.transform = 'translateY(-1px)';
              btn.style.boxShadow = '0 8px 18px rgba(14, 106, 78, 0.45)';
              btn.style.opacity = '1';
          });
          btn.addEventListener('mouseout', () => {
              btn.style.transform = '';
              btn.style.boxShadow = '0 4px 12px rgba(14, 106, 78, 0.35)';
              btn.style.opacity = '0.95';
          });
          btn.addEventListener('click', function () {
              // Decide whether to OPEN or CLOSE the sidebar based on its
              // current width, then click the matching native control.
              const sb = hostDoc.querySelector('section[data-testid="stSidebar"]');
              const isOpen = sb && sb.getBoundingClientRect().width > 100;
              const tryClick = (sel) => {
                  const el = hostDoc.querySelector(sel);
                  if (el) { el.click(); return true; }
                  return false;
              };
              if (isOpen) {
                  tryClick('section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] button')
                    || tryClick('section[data-testid="stSidebar"] button[kind="header"]')
                    || tryClick('section[data-testid="stSidebar"] button[aria-label*="sidebar" i]');
              } else {
                  tryClick('[data-testid="stSidebarCollapsedControl"] button')
                    || tryClick('[data-testid="collapsedControl"] button')
                    || tryClick('button[data-testid="stExpandSidebarButton"]')
                    || tryClick('button[kind="header"][aria-label*="sidebar" i]');
              }
          });
          hostDoc.body.appendChild(btn);
      }

      // CRITICAL: never set display:none on this button. Earlier versions
      // tried to hide it when the sidebar was open and got latched into a
      // permanently-hidden state because Streamlit fires the mutation
      // observer mid-CSS-transition, when width is still ~250px. Just
      // always show it — a small overlap when the sidebar is open is fine.
      btn.style.display = 'flex';
  })();
</script>
"""


def inject_floating_menu_button() -> None:
    """Render a permanent ☰ button at the top-left to re-open the sidebar.

    The iframe itself is invisible; only the fixed-position button is shown.
    """
    # Local import keeps this an optional dependency if streamlit is upgraded.
    from streamlit.components.v1 import html
    html(_FLOATING_MENU_HTML, height=0)


def section_title(text: str, hint: str | None = None) -> None:
    """Render a styled section heading with the pink accent bar."""
    hint_html = f'<div class="section-title__hint">{hint}</div>' if hint else ""
    st.markdown(
        f"""
        <div class="section-title">
            <div class="section-title__left">
                <div class="section-title__bar"></div>
                <div class="section-title__text">{text}</div>
            </div>
            {hint_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def alert(kind: str, body_html: str, icon: str = "ℹ️") -> None:
    """Render an info/ok/warn/err alert."""
    st.markdown(
        f"""
        <div class="alert alert--{kind}">
            <div class="alert__icon">{icon}</div>
            <div class="alert__body">{body_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
