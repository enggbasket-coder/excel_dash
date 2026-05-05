# Excel VBA Dashboard

A modern Streamlit control surface for running Excel macros, drafting EPC / contract correspondence, and conducting site audits — all in one place.

![Built with Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-0E6A4E)
![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue)

## Tabs

- **Macro** — KPI strip, Excel-status detector, six quick-preset macros, custom macro runner, recent-runs log.
- **Letters** — eight EPC templates (RFI, EOT, Variation, NCR, Progress, Hand-over, Payment, Closure), a composer, and a saved-drafts list.
- **Audit** — observation form, four-category compliance checklist with progress bars, photo-record gallery, and a kanban-style corrective-actions board.
- **Help** — quick start, important notes, FAQs, troubleshooting, and shortcuts.

## Run locally

```bash
git clone https://github.com/<your-username>/excel-vba-dashboard.git
cd excel-vba-dashboard
bash run.sh
```

The script creates a `.venv`, installs dependencies, and launches Streamlit on **http://localhost:8502**.

## Run on Streamlit Community Cloud

Push this repo to GitHub, then connect it at <https://share.streamlit.io>. Set the entry point to `app.py` — that's it.

> **Cloud limitation:** the Macro tab talks to a *local* Excel via AppleScript (macOS) or pywin32 (Windows). On the Linux servers Streamlit Cloud uses, neither is available, so the Macro tab will show "No Excel session detected" — that is expected. The Letters, Audit, and Help tabs work fully on the cloud.

## Project structure

```
.
├── app.py                  # Entry point — page config, theme, sidebar, tab routing
├── ui/
│   ├── theme.py            # Mint + forest-green design system, CSS, floating menu
│   └── sidebar.py          # Brand block, tab navigation
├── core/
│   └── excel_bridge.py     # Cross-platform Excel bridge (AppleScript / pywin32)
├── tabs/
│   ├── macro.py            # KPIs, status, presets, custom runner, log
│   ├── letters.py          # Templates, composer, drafts
│   ├── audit.py            # Observations, compliance, photos, actions
│   └── help.py             # Quick start, notes, FAQ, troubleshooting
├── .streamlit/config.toml  # Theme & server settings
├── requirements.txt        # streamlit, pywin32 (Windows-only)
└── run.sh                  # One-command local launcher
```

## License

MIT — do whatever you want.
