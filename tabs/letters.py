"""
Letters tab — EPC/contract/project correspondence workspace.

Layout:
    1. KPI strip  (Total drafts / This month / Templates / Sent)
    2. Templates  (card grid — RFI, EOT, Variation, NCR, Closure, Hand-over…)
    3. Compose    (editor card with metadata fields + body)
    4. Recent     (saved drafts list with status pills)
"""

from __future__ import annotations

from datetime import datetime

import streamlit as st

from ui.theme import section_title, alert


TEMPLATES = [
    {"key": "RFI",        "icon": "❓", "name": "Request for Information",
     "desc": "Standard RFI with reference to drawing/spec clause."},
    {"key": "EOT",        "icon": "⏰", "name": "Extension of Time",
     "desc": "Notice & particulars under contract clause."},
    {"key": "VARIATION",  "icon": "🔧", "name": "Variation Order",
     "desc": "Scope-change request with cost & time impact."},
    {"key": "NCR",        "icon": "⚠️", "name": "Non-Conformance Report",
     "desc": "Quality NCR with corrective action."},
    {"key": "PROGRESS",   "icon": "📈", "name": "Progress Report",
     "desc": "Weekly/monthly progress reporting cover letter."},
    {"key": "HANDOVER",   "icon": "🤝", "name": "Hand-Over Letter",
     "desc": "Section/snag-list completion hand-over."},
    {"key": "PAYMENT",    "icon": "💳", "name": "Payment Certificate Cover",
     "desc": "IPC/RA-bill submission cover note."},
    {"key": "CLOSURE",    "icon": "✅", "name": "Project Closure",
     "desc": "Final closure notification & defect-liability."},
]


_TEMPLATE_BODIES = {
    "RFI":
        "Dear Sir/Madam,\n\nWith reference to the above subject, we kindly request "
        "clarification on the following:\n\n1. \n2. \n\nAn early response would be "
        "appreciated to avoid impact on the construction schedule.\n\nThanking you,\n",
    "EOT":
        "Dear Sir/Madam,\n\nIn accordance with Clause [__] of the Contract, we hereby "
        "submit our Notice of Delay and request for Extension of Time on the following "
        "grounds:\n\n1. Cause of delay: \n2. Period claimed: \n3. Particulars enclosed: "
        "Annexure-A\n\nWe request your kind concurrence at the earliest.\n\nThanking you,\n",
    "VARIATION":
        "Dear Sir/Madam,\n\nWe wish to bring to your notice the following variation in scope:\n\n"
        "Description of variation:\nReason:\nCost impact (INR):\nTime impact (days):\n\n"
        "Kindly issue a Variation Order to enable us to proceed.\n\nThanking you,\n",
    "NCR":
        "Dear Sir/Madam,\n\nDuring our routine inspection on [date], the following "
        "non-conformance was observed:\n\nObservation:\nLocation:\nReference clause:\n"
        "Proposed corrective action:\n\nWe await your instructions for closure.\n\nThanking you,\n",
    "PROGRESS":
        "Dear Sir/Madam,\n\nPlease find enclosed the [Weekly/Monthly] Progress Report for "
        "the period [__] to [__]. Key highlights:\n\n• Physical progress: __ %\n"
        "• Financial progress: __ %\n• Critical issues: \n\nLooking forward to your "
        "feedback at the next review meeting.\n\nThanking you,\n",
    "HANDOVER":
        "Dear Sir/Madam,\n\nWe are pleased to hand over the following section/works for "
        "your acceptance:\n\nSection/Works: \nDate of completion: \nSnag-list status: \n\n"
        "Kindly arrange joint inspection at your earliest convenience.\n\nThanking you,\n",
    "PAYMENT":
        "Dear Sir/Madam,\n\nWe submit herewith our Interim Payment Certificate / RA-Bill "
        "No. [__] for the period [__] to [__], for an amount of INR [__].\n\nSupporting "
        "documents are enclosed for your verification & onward processing.\n\nThanking you,\n",
    "CLOSURE":
        "Dear Sir/Madam,\n\nWith reference to the captioned subject, we confirm that "
        "all contractual obligations under the project have been fulfilled. We request "
        "you to kindly issue the Project Completion Certificate and release the retention "
        "amount as per the contract terms.\n\nThanking you,\n",
}


# ───────────────────────────────────────────────────────────────
def _kpis() -> None:
    drafts = st.session_state.letter_drafts
    total = len(drafts)
    this_month = sum(1 for d in drafts
                     if d["created"].month == datetime.now().month
                     and d["created"].year == datetime.now().year)
    sent = sum(1 for d in drafts if d.get("status") == "Sent")

    cols = st.columns(4, gap="medium")
    kpis = [
        ("📝", "Total Drafts",   str(total),       "all-time"),
        ("📅", "This Month",     str(this_month),  "in May 2026"),
        ("📂", "Templates",      str(len(TEMPLATES)), "ready to use"),
        ("📤", "Sent",           str(sent),        "marked sent"),
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
def _templates_grid() -> None:
    section_title("Letter Templates",
                  hint="Pick a template to pre-fill the editor below")
    rows = [TEMPLATES[i:i + 4] for i in range(0, len(TEMPLATES), 4)]
    for row in rows:
        cols = st.columns(4, gap="medium")
        for col, t in zip(cols, row):
            with col:
                st.markdown(
                    f"""
                    <div class="macro-tile">
                        <div class="macro-tile__icon">{t['icon']}</div>
                        <div class="macro-tile__name">{t['name']}</div>
                        <div class="macro-tile__desc">{t['desc']}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
                if st.button(f"Use", key=f"tpl_{t['key']}", use_container_width=True):
                    st.session_state.letter_template_key = t["key"]
                    st.session_state.letter_template_name = t["name"]
                    st.session_state.letter_body = _TEMPLATE_BODIES.get(t["key"], "")
                    st.rerun()


# ───────────────────────────────────────────────────────────────
def _compose_card() -> None:
    section_title("Compose Letter",
                  hint=f"Editing: {st.session_state.get('letter_template_name', 'Blank letter')}")

    with st.container():
        c1, c2, c3 = st.columns(3, gap="medium")
        with c1:
            ref = st.text_input(
                "Letter reference",
                key="letter_ref",
                placeholder="e.g. CONT/RFI/2026/045",
            )
        with c2:
            recipient = st.text_input(
                "To (name & designation)",
                key="letter_to",
                placeholder="e.g. Mr. A. Sharma, Project Manager",
            )
        with c3:
            project = st.text_input(
                "Project / Contract",
                key="letter_project",
                placeholder="e.g. NH-44 Rehabilitation Pkg-3",
            )

        subject = st.text_input(
            "Subject",
            key="letter_subject",
            placeholder="Subject of the letter",
        )

        body = st.text_area(
            "Body",
            key="letter_body",
            height=260,
            placeholder="Compose the letter body here…",
        )

        b1, b2, b3, _ = st.columns([1.2, 1.2, 1.2, 4], gap="small")
        with b1:
            save_clicked = st.button("💾 Save Draft", use_container_width=True, type="primary")
        with b2:
            mark_sent = st.button("📤 Mark Sent", use_container_width=True)
        with b3:
            clear = st.button("🧹 Clear", use_container_width=True)

        if save_clicked:
            if not subject.strip():
                alert("warn", "Please enter a <b>subject</b> before saving.", icon="⚠️")
            else:
                st.session_state.letter_drafts.append({
                    "ref": ref or f"DRAFT-{len(st.session_state.letter_drafts)+1:03d}",
                    "to": recipient,
                    "project": project,
                    "subject": subject,
                    "body": body,
                    "template": st.session_state.get("letter_template_name", "Blank"),
                    "status": "Draft",
                    "created": datetime.now(),
                })
                alert("ok", f"Draft <b>{subject}</b> saved.", icon="✅")

        if mark_sent and st.session_state.letter_drafts:
            st.session_state.letter_drafts[-1]["status"] = "Sent"
            alert("ok", "Latest draft marked as <b>Sent</b>.", icon="📤")

        if clear:
            for k in ["letter_ref", "letter_to", "letter_project",
                      "letter_subject", "letter_body",
                      "letter_template_key", "letter_template_name"]:
                st.session_state.pop(k, None)
            st.rerun()


# ───────────────────────────────────────────────────────────────
def _recent_drafts() -> None:
    section_title("Recent Letters",
                  hint=f"{len(st.session_state.letter_drafts)} saved")

    drafts = st.session_state.letter_drafts
    if not drafts:
        st.markdown(
            """
            <div class="card" style="text-align:center;color:var(--ink-500);padding:32px;">
                <div style="font-size:28px;margin-bottom:6px;">📭</div>
                No saved drafts yet. Compose one above and click <b>Save Draft</b>.
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    for d in reversed(drafts[-12:]):
        is_sent = d["status"] == "Sent"
        badge_bg = "#DCFCE7" if is_sent else "#FEF3C7"
        badge_fg = "#166534" if is_sent else "#92400E"
        st.markdown(
            f"""
            <div class="card" style="margin-bottom:10px;">
                <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;">
                    <div style="flex:1;min-width:0;">
                        <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;">
                            <span style="font-weight:700;color:var(--ink-900);">{d['subject']}</span>
                            <span style="background:{badge_bg};color:{badge_fg};padding:3px 10px;
                                  border-radius:999px;font-weight:600;font-size:11.5px;">
                                {d['status'].upper()}
                            </span>
                        </div>
                        <div style="color:var(--ink-500);font-size:12.5px;margin-top:4px;">
                            {d['ref']} · To: {d['to'] or '—'} · {d['project'] or '—'} ·
                            <i>{d['template']}</i> ·
                            {d['created'].strftime('%d %b %Y, %H:%M')}
                        </div>
                    </div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# ───────────────────────────────────────────────────────────────
def render_letters_tab() -> None:
    _kpis()
    st.write("")
    _templates_grid()
    st.write("")
    _compose_card()
    st.write("")
    _recent_drafts()
