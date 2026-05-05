"""
Audit tab — site observations, compliance checks, photo records,
and corrective-action tracking.

Sections:
    1. KPI strip (Open / In progress / Closed / Photos)
    2. New Observation (form card)
    3. Compliance Checklist (categorical checklist with progress)
    4. Photo Records (uploader + gallery)
    5. Corrective Actions (kanban-style status board)
"""

from __future__ import annotations

from datetime import datetime, date
from pathlib import Path

import streamlit as st

from ui.theme import section_title, alert


COMPLIANCE_CATEGORIES = {
    "Safety (HSE)": [
        "PPE compliance on site",
        "First-aid kit & emergency contacts displayed",
        "Toolbox talk conducted today",
        "Fire extinguishers serviced & accessible",
        "Working-at-height permits in place",
    ],
    "Quality (QC)": [
        "Material test certificates available",
        "Mix-design approved & followed",
        "Curing as per IS code",
        "Reinforcement cover blocks used",
        "Concrete cube samples taken & labelled",
    ],
    "Documentation": [
        "Daily Progress Report filed",
        "Drawings revision-controlled",
        "RFIs logged & tracked",
        "NCRs closed within 7 days",
        "Hindrance register updated",
    ],
    "Environmental": [
        "Dust suppression in place",
        "Spoil disposal as per consent",
        "Diesel storage bunded",
        "Noise levels within limits",
        "Tree-cutting permits valid",
    ],
}


# ───────────────────────────────────────────────────────────────
def _ensure_state():
    if "audit_compliance" not in st.session_state:
        st.session_state.audit_compliance = {
            cat: {item: False for item in items}
            for cat, items in COMPLIANCE_CATEGORIES.items()
        }
    if "audit_photos" not in st.session_state:
        st.session_state.audit_photos = []   # list of (filename, bytes)
    if "audit_actions" not in st.session_state:
        st.session_state.audit_actions = []


# ───────────────────────────────────────────────────────────────
def _kpis():
    obs = st.session_state.audit_observations
    actions = st.session_state.audit_actions
    open_n   = sum(1 for o in obs if o["status"] == "Open")
    prog_n   = sum(1 for o in obs if o["status"] == "In Progress")
    closed_n = sum(1 for o in obs if o["status"] == "Closed")
    photos_n = len(st.session_state.audit_photos)

    cols = st.columns(4, gap="medium")
    kpis = [
        ("🔓", "Open Observations",   str(open_n),    "needs action"),
        ("🛠", "In Progress",          str(prog_n),    "being worked"),
        ("🔒", "Closed",               str(closed_n),  "resolved"),
        ("📸", "Photo Records",        str(photos_n),  "uploaded"),
    ]
    for col, (icon, label, value, delta) in zip(cols, kpis):
        with col:
            st.markdown(
                f"""
                <div class="kpi">
                    <div class="kpi__icon">{icon}</div>
                    <div class="kpi__label">{label}</div>
                    <div class="kpi__value">{value}</div>
                    <div class="kpi__delta">{delta}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# ───────────────────────────────────────────────────────────────
def _new_observation():
    section_title("New Audit Observation",
                  hint="Log site / quality / safety observations")

    c1, c2, c3 = st.columns(3, gap="medium")
    with c1:
        category = st.selectbox(
            "Category",
            options=["Safety", "Quality", "Documentation", "Environmental", "Other"],
            key="obs_category",
        )
    with c2:
        severity = st.selectbox(
            "Severity",
            options=["Low", "Medium", "High", "Critical"],
            key="obs_severity",
        )
    with c3:
        location = st.text_input(
            "Location",
            placeholder="e.g. Pier P-12, Ch. 4+200",
            key="obs_location",
        )

    title = st.text_input(
        "Observation title", placeholder="Short headline of the observation",
        key="obs_title",
    )
    detail = st.text_area(
        "Description / Recommended action",
        height=110,
        placeholder="Describe what was observed and the corrective action recommended…",
        key="obs_detail",
    )

    bcol1, bcol2, _ = st.columns([1.4, 1.4, 4])
    with bcol1:
        save = st.button("➕ Add Observation", type="primary", use_container_width=True)
    with bcol2:
        clear = st.button("🧹 Clear form", use_container_width=True)

    if save:
        if not title.strip():
            alert("warn", "Please enter an <b>observation title</b>.", icon="⚠️")
        else:
            st.session_state.audit_observations.append({
                "id": f"OBS-{len(st.session_state.audit_observations)+1:03d}",
                "title": title.strip(),
                "category": category,
                "severity": severity,
                "location": location,
                "detail": detail,
                "status": "Open",
                "created": datetime.now(),
            })
            alert("ok", f"Observation <b>{title}</b> recorded.", icon="✅")

    if clear:
        for k in ["obs_title", "obs_detail", "obs_location"]:
            st.session_state.pop(k, None)
        st.rerun()


# ───────────────────────────────────────────────────────────────
def _compliance_checklist():
    section_title("Compliance Checklist",
                  hint="Tick items confirmed during today's audit")

    state = st.session_state.audit_compliance

    for cat, items in COMPLIANCE_CATEGORIES.items():
        ticked = sum(1 for item in items if state[cat][item])
        pct = int(ticked / len(items) * 100)
        bar_color = ("var(--good)" if pct == 100
                     else "var(--pink-500)" if pct >= 60
                     else "var(--warn)")

        st.markdown(
            f"""
            <div class="card">
                <div style="display:flex;justify-content:space-between;align-items:center;
                            margin-bottom:10px;">
                    <div style="font-weight:700;color:var(--ink-900);">{cat}</div>
                    <div style="color:var(--ink-500);font-size:12.5px;font-weight:600;">
                        {ticked}/{len(items)} · {pct}%
                    </div>
                </div>
                <div style="height:8px;background:var(--ink-100);border-radius:99px;overflow:hidden;
                            margin-bottom:10px;">
                    <div style="height:100%;width:{pct}%;background:{bar_color};
                                transition:width .3s ease;"></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        with st.expander(f"View {cat.lower()} items", expanded=False):
            for item in items:
                state[cat][item] = st.checkbox(
                    item,
                    value=state[cat][item],
                    key=f"chk_{cat}_{item}",
                )


# ───────────────────────────────────────────────────────────────
def _photo_records():
    section_title("Photo Records",
                  hint="Attach geo-tagged site photos for evidence")

    files = st.file_uploader(
        "Drop photos here or click to browse",
        type=["png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key="audit_uploader",
    )
    if files:
        for f in files:
            # Avoid re-adding the same file on rerun by checking name+size
            sig = (f.name, f.size)
            if sig not in [(p["name"], p["size"]) for p in st.session_state.audit_photos]:
                st.session_state.audit_photos.append({
                    "name": f.name,
                    "size": f.size,
                    "data": f.read(),
                    "uploaded": datetime.now(),
                })

    photos = st.session_state.audit_photos
    if not photos:
        st.markdown(
            """
            <div class="card" style="text-align:center;color:var(--ink-500);padding:24px;">
                <div style="font-size:24px;margin-bottom:4px;">🖼</div>
                No photos uploaded yet.
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    # Gallery — 4 columns
    cols = st.columns(4, gap="small")
    for i, photo in enumerate(photos[-12:]):
        with cols[i % 4]:
            st.image(photo["data"], caption=photo["name"], use_container_width=True)


# ───────────────────────────────────────────────────────────────
def _corrective_actions():
    section_title("Corrective Actions",
                  hint="Track the status of every observation through to closure")

    obs = st.session_state.audit_observations
    if not obs:
        st.markdown(
            """
            <div class="card" style="text-align:center;color:var(--ink-500);padding:24px;">
                <div style="font-size:24px;margin-bottom:4px;">📋</div>
                No observations logged yet — add one above to start tracking.
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    # Three-column kanban: Open / In Progress / Closed
    cols = st.columns(3, gap="medium")
    columns_meta = [
        ("Open",        "🔴", "#FEE2E2", "#991B1B"),
        ("In Progress", "🟡", "#FEF3C7", "#92400E"),
        ("Closed",      "🟢", "#DCFCE7", "#166534"),
    ]
    sev_color = {"Low": "#94A3B8", "Medium": "#3B82F6",
                 "High": "#F59E0B", "Critical": "#DC2626"}

    for col, (status_key, dot, bg, fg) in zip(cols, columns_meta):
        items = [o for o in obs if o["status"] == status_key]
        with col:
            st.markdown(
                f"""
                <div style="display:flex;justify-content:space-between;align-items:center;
                            padding:8px 4px 12px 4px;">
                    <div style="font-weight:700;color:var(--ink-900);font-size:13.5px;">
                        {dot} {status_key}
                    </div>
                    <div style="background:{bg};color:{fg};padding:2px 10px;border-radius:999px;
                                font-weight:700;font-size:11.5px;">{len(items)}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if not items:
                st.markdown(
                    f"""
                    <div style="border:1.5px dashed var(--ink-100);border-radius:14px;
                                padding:18px;text-align:center;color:var(--ink-300);
                                font-size:12.5px;">No items</div>
                    """,
                    unsafe_allow_html=True,
                )
                continue
            for o in items:
                sev_chip = sev_color.get(o["severity"], "#94A3B8")
                st.markdown(
                    f"""
                    <div class="card" style="margin-bottom:10px;padding:14px 16px;">
                        <div style="display:flex;justify-content:space-between;
                                    align-items:flex-start;gap:8px;">
                            <div style="font-weight:700;color:var(--ink-900);font-size:13.5px;">
                                {o['title']}
                            </div>
                            <span style="background:{sev_chip}1A;color:{sev_chip};
                                  padding:2px 8px;border-radius:999px;font-weight:600;
                                  font-size:11px;">{o['severity']}</span>
                        </div>
                        <div style="color:var(--ink-500);font-size:11.5px;margin-top:6px;">
                            {o['id']} · {o['category']} · {o['location'] or '—'}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                # Action buttons (advance / regress)
                bcols = st.columns(2)
                if status_key != "Closed":
                    with bcols[0]:
                        if st.button("→ Advance", key=f"adv_{o['id']}",
                                     use_container_width=True):
                            o["status"] = ("In Progress" if status_key == "Open"
                                           else "Closed")
                            st.rerun()
                if status_key != "Open":
                    with bcols[1]:
                        if st.button("← Reopen", key=f"reo_{o['id']}",
                                     use_container_width=True):
                            o["status"] = ("In Progress" if status_key == "Closed"
                                           else "Open")
                            st.rerun()


# ───────────────────────────────────────────────────────────────
def render_audit_tab() -> None:
    _ensure_state()
    _kpis()
    st.write("")
    _new_observation()
    st.write("")
    _compliance_checklist()
    st.write("")
    _photo_records()
    st.write("")
    _corrective_actions()
