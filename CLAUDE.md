# Excel VBA Dashboard — project notes

Streamlit dashboard (port 8502) that drives the user's **live Excel on macOS** to run
VBA macros on highway-construction measurement workbooks (SPS bills: S.G, GSB, WMM,
DBM, BC layer sheets with From/To chainage + Side columns).

Start: `./run.sh` → http://localhost:8502 (creates `.venv` on first run).

## Tabs
- **StripChart Plot** (`tabs/stripchart.py`) — user picks an open workbook, chainage
  range, up to 7 layer slots (first = bottom layer). Dashboard writes a hidden
  `_SC_Cfg` sheet into that workbook (B1 from_m, B2 to_m, B3 interval, B4 assume_km,
  B5:B11 layer sheets, C5 strip-sheet name), saves, then runs `StripChart_Run` via
  osascript. Macro logs to `_SC_Cfg` column Z; dashboard reads the log.
- **Macro** (`tabs/macro.py`) — preset buttons `FindOverlapsAllSheets`, `RFI_dupli`
  run via `core/excel_bridge.py` on the ACTIVE workbook.
- `core/stripchart_engine.py` — pure-Python strip chart engine (for a future
  upload/online version; not used by the live tab).

## Macro workbooks (not in git — `*.xlsm` is gitignored, copy them across manually)
- Strip chart: `StripChart_macros Final working V2.xlsm` (source: `vba/StripChart.bas`)
- Overlap + RFI duplicate: `Overlap_dup_tool.xlsm` (`FindOverlapsAllSheets`,
  `RFI_dupli`, `Payprocedure`; source of RFI audit: `vba/RFI_Duplicate_Audit.bas`)
- An OLD `BuildStripChart.bas` (input-box version) exists elsewhere — do not import it.

## Hard-won gotchas (macOS Excel)
- AppleScript must be `run VB Macro "Name"` — the macro name is a DIRECT parameter
  (per Excel.sdef). `run VB Macro Macro Name "X"` fails to compile (-2741/-2740).
- Wrap in `with timeout of 1800 seconds` or long runs fail with -1712.
- xlwings `app.macro()` silently does nothing on Mac — don't use it there.
- `StripChart_RunImpl` finds the data workbook by scanning open books for `_SC_Cfg`
  (skipping `ThisWorkbook`), so the .xlsm can have any filename.
- Data columns are auto-detected from headers "From"/"To"/"Side" (SPS-12: E/F/G,
  SPS-08: F/G/I). Data rows start at first numeric cell in column A.
- Ranges are half-open [From, To) — adjacent rows sharing a boundary are NOT overlaps.
- Chainage parsing: "152+780", "152.78" (km) and 152780 (≥1000 ⇒ metres).
- Sheet names often have trailing spaces ("S.G ", "WMM ") — pick from dropdown.
- Strip sheet name must differ from every data sheet name (macro aborts otherwise).
- VBA: helper params that loop on themselves need `ByVal` (ColLetter bug zeroed cols).
- Chart range must match the project: SPS-08 = 99+700→119+015; SPS-12/SPS-13 =
  152+490→166+700. Wrong range ⇒ "total_painted=0" (dashboard now warns).
- `Overlap_dup_tool.xlsm` contains an LLM module that fails to compile on Mac
  ("Automation type not supported") and blocks all its macros — remove that module.

## Audit output (StripChart)
Overlap cells painted red, upper-layer-without-lower-layer cells purple; per-layer
remarks in the column two to the right of the chart (green "OK" / red issue text).
Blinking was tried and removed at the user's request.

## Open ideas
- Online version: port `FindOverlapsAllSheets` and `RFI_dupli` to Python and switch
  StripChart tab to upload/download using `core/stripchart_engine.py`.
