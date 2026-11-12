# stripchart.spec
# -----------------------------------------------------------------------
# Cross-platform PyInstaller spec for StripChart Dashboard (Excel_dash_UI)
#
# Build on Mac:     pyinstaller stripchart.spec          (from Excel_dash_UI/)
# Build on Windows: pyinstaller stripchart.spec          (from Excel_dash_UI\)
#
# Output
#   Mac     → dist/StripChartDashboard.app   (double-click to launch)
#   Windows → dist/StripChartDashboard/StripChartDashboard.exe
# -----------------------------------------------------------------------

import sys
from PyInstaller.utils.hooks import collect_all

# --- Collect Streamlit (has many internal data files and hidden imports)
st_datas, st_bins, st_hidden = collect_all("streamlit")

# --- Altair (Streamlit's default chart backend — optional)
try:
    al_datas, al_bins, al_hidden = collect_all("altair")
except Exception:
    al_datas, al_bins, al_hidden = [], [], []

# --- pydeck (Streamlit map widget — optional)
try:
    pd_datas, pd_bins, pd_hidden = collect_all("pydeck")
except Exception:
    pd_datas, pd_bins, pd_hidden = [], [], []

a = Analysis(
    ["launcher.py"],
    # pathex must include this folder so PyInstaller can resolve
    # local packages (tabs, ui, core) during analysis.
    pathex=["."],
    binaries=st_bins + al_bins + pd_bins,
    datas=[
        # ── App entry point ──────────────────────────────────────────
        # Streamlit runs app.py as a file at runtime — it must be present
        # as a .py source file inside the bundle, not only as bytecode.
        ("app.py", "."),

        # ── Local packages ────────────────────────────────────────────
        # Include the full source trees so `app.py` can import them.
        # PyInstaller places them under sys._MEIPASS which is already
        # on sys.path, so `from tabs.stripchart import …` just works.
        ("tabs", "tabs"),
        ("ui",   "ui"),
        ("core", "core"),

        # ── Macro workbook ────────────────────────────────────────────
        # Bundled alongside the app; launcher.py copies it to
        # ~/Documents/ExcelDashboard/ on first run and opens it in Excel.
        ("../Excel_dashboard/StripChart_macros.xlsm", "."),

        # ── Streamlit internal assets ─────────────────────────────────
        *st_datas,
        *al_datas,
        *pd_datas,
    ],
    hiddenimports=[
        *st_hidden,
        *al_hidden,
        *pd_hidden,

        # ── Core data stack ───────────────────────────────────────────
        "xlwings",
        "pandas",
        "openpyxl",
        "pyarrow",
        "pyarrow.vendored.version",
        "requests",

        # ── Streamlit internals (sometimes missed by collect_all) ─────
        "streamlit.web.cli",
        "streamlit.web.server",
        "streamlit.runtime.scriptrunner",

        # ── Windows: COM / Excel automation ───────────────────────────
        "win32com",
        "win32com.client",
        "comtypes",
        "comtypes.client",

        # ── macOS: AppleScript bridge used by xlwings ─────────────────
        "appscript",
        "mactypes",

        # ── Image handling (Streamlit optional) ───────────────────────
        "PIL",
        "PIL.Image",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        # Not used — saves ~50 MB from the bundle
        "tkinter",
        "matplotlib",
        "scipy",
        "sklearn",
        "notebook",
        "IPython",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="StripChartDashboard",
    debug=False,
    strip=False,
    upx=True,
    # console=True keeps a terminal window visible — useful while testing.
    # Change to False on Windows for a cleaner distribution once all is OK.
    console=True,
    target_arch=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name="StripChartDashboard",
)

# macOS only — wrap COLLECT output into a clickable .app bundle
if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="StripChartDashboard.app",
        icon=None,
        bundle_identifier="com.stripchart.dashboard",
        info_plist={
            "NSHighResolutionCapable": True,
            "CFBundleShortVersionString": "1.0.0",
            # Required so macOS allows this app to send AppleEvents to Excel
            "NSAppleEventsUsageDescription": (
                "This app controls Microsoft Excel to run "
                "strip-chart macros on your workbook."
            ),
        },
    )
