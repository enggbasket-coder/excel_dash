#!/bin/bash
# build_mac.sh — builds StripChartDashboard.app on macOS
#
# Run from the Excel_dash_UI/ folder:
#   bash build_mac.sh
#
# Output: dist/StripChartDashboard.app
# -------------------------------------------------------
set -e

echo "==================================================="
echo "  StripChart Dashboard — macOS build"
echo "==================================================="

# ── 1. Locate StripChart_macros.xlsm ────────────────────
XLSM="../Excel_dashboard/StripChart_macros.xlsm"
if [ ! -f "$XLSM" ]; then
    echo ""
    echo "ERROR: StripChart_macros.xlsm not found."
    echo "Expected location: $XLSM"
    echo "Make sure the Excel_dashboard folder is next to Excel_dash_UI/."
    exit 1
fi
echo "Found: $XLSM"

# ── 2. Install / upgrade build dependencies ─────────────
echo ""
echo "Installing build dependencies..."
pip install pyinstaller --quiet
pip install -r requirements.txt --quiet

# ── 3. Clean previous builds ────────────────────────────
echo "Cleaning previous builds..."
rm -rf build dist

# ── 4. Build ────────────────────────────────────────────
echo ""
echo "Building... (this takes 2–5 minutes)"
pyinstaller stripchart.spec

# ── 5. Done ─────────────────────────────────────────────
echo ""
echo "==================================================="
echo "  Build complete!"
echo ""
echo "  App:  dist/StripChartDashboard.app"
echo ""
echo "  To distribute to Mac users:"
echo "    cd dist"
echo "    zip -r StripChartDashboard_mac.zip StripChartDashboard.app"
echo "    # Share StripChartDashboard_mac.zip"
echo ""
echo "  User instructions:"
echo "    1. Unzip → double-click StripChartDashboard.app"
echo "    2. If macOS blocks it: right-click → Open → 'Open' in the dialog"
echo "    3. App auto-opens StripChart_macros.xlsm in Excel"
echo "    4. Open your data workbook in Excel"
echo "    5. Use the StripChart tab in the browser dashboard"
echo "==================================================="
