"""
Entry point for the PyInstaller bundle — StripChart Dashboard.

When frozen (double-clicked by the user):
  - Copies StripChart_macros.xlsm from the bundle to
    ~/Documents/ExcelDashboard/  (writable, Excel can always open from here)
  - Auto-opens the xlsm in Excel on Mac so macros are available immediately.
  - Opens the browser after Streamlit starts (~2.5 s).
  - Launches Streamlit on port 8501.

Run manually in development (no PyInstaller needed):
  cd Excel_dash_UI
  python launcher.py
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
import threading
import time
import webbrowser

IS_MAC = platform.system() == "Darwin"
IS_WIN = platform.system() == "Windows"

XLSM_NAME = "StripChart_macros.xlsm"


# ─────────────────────────────────────────────────────────────
#  Path helpers
# ─────────────────────────────────────────────────────────────

def _base_dir() -> str:
    """Root of the bundle (_MEIPASS when frozen, script dir in dev)."""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS  # type: ignore[attr-defined]
    return os.path.dirname(os.path.abspath(__file__))


def _app_script() -> str:
    """Absolute path to app.py inside the bundle (or dev folder)."""
    return os.path.join(_base_dir(), "app.py")


def _user_xlsm_dir() -> str:
    """
    Writable folder in the user's home directory.
    Excel has no sandbox restrictions reading from ~/Documents/.
    """
    return os.path.join(os.path.expanduser("~"), "Documents", "ExcelDashboard")


# ─────────────────────────────────────────────────────────────
#  XLSM setup
# ─────────────────────────────────────────────────────────────

def _setup_xlsm() -> str | None:
    """
    Copy StripChart_macros.xlsm from the bundle to ~/Documents/ExcelDashboard/
    and open it in Excel (Mac) so the VBA macros are ready.

    Returns the destination path on success, None if the xlsm is not bundled
    (e.g. when running in dev mode without the file alongside launcher.py).
    """
    bundled_xlsm = os.path.join(_base_dir(), XLSM_NAME)
    if not os.path.exists(bundled_xlsm):
        print(f"[Setup] {XLSM_NAME} not found in bundle — "
              "open it manually in Excel before running a strip chart.")
        return None

    dest_dir = _user_xlsm_dir()
    os.makedirs(dest_dir, exist_ok=True)
    dest_path = os.path.join(dest_dir, XLSM_NAME)

    # Only copy when the bundled file is newer (or destination is absent).
    needs_copy = (
        not os.path.exists(dest_path)
        or os.path.getmtime(bundled_xlsm) > os.path.getmtime(dest_path)
    )
    if needs_copy:
        shutil.copy2(bundled_xlsm, dest_path)
        print(f"[Setup] Copied {XLSM_NAME}  →  {dest_path}")
    else:
        print(f"[Setup] {XLSM_NAME} already up to date at {dest_path}")

    # Expose location to the app (optional — the VBA discovery also works via xlwings)
    os.environ["STRIPCHART_XLSM_DIR"] = dest_dir

    # Auto-open in Excel on Mac so the macro workbook is ready immediately.
    if IS_MAC:
        try:
            subprocess.Popen(["open", "-a", "Microsoft Excel", dest_path])
            print(f"[Setup] Opened {XLSM_NAME} in Excel.")
        except Exception as exc:
            print(f"[Setup] Could not auto-open Excel: {exc}")
            print(f"[Setup] Please open {dest_path} in Excel manually.")

    # Windows: user opens their data workbook themselves; the macro workbook
    # path is shown in the dashboard's StripChart tab instructions.
    if IS_WIN:
        print(f"[Setup] Windows: open {dest_path} in Excel before running macros.")

    return dest_path


# ─────────────────────────────────────────────────────────────
#  Browser helper
# ─────────────────────────────────────────────────────────────

def _open_browser(url: str, delay: float = 2.5) -> None:
    """Open the browser after a short delay (Streamlit takes ~2 s to start)."""
    def _go() -> None:
        time.sleep(delay)
        webbrowser.open(url)
    threading.Thread(target=_go, daemon=True).start()


# ─────────────────────────────────────────────────────────────
#  Main
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    base = _base_dir()
    os.environ["EXCEL_DASHBOARD_BASE"] = base

    # Step 1 — deploy the macro workbook
    _setup_xlsm()

    # Step 2 — open the browser once Streamlit is up
    _open_browser("http://localhost:8501")

    # Step 3 — launch Streamlit (blocking call)
    import streamlit.web.cli as stcli  # noqa: PLC0415

    sys.argv = [
        "streamlit", "run", _app_script(),
        "--server.port=8501",
        "--server.headless=false",
        "--browser.gatherUsageStats=false",
        "--server.fileWatcherType=none",   # no hot-reload inside the bundle
    ]
    sys.exit(stcli.main())
