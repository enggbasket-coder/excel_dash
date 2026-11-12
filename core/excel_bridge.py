"""
Cross-platform bridge to a *running* Excel session.

macOS  → AppleScript (osascript) talking to "Microsoft Excel"
Windows → win32com.client via pywin32 (lazy-imported so non-Windows machines don't fail)

Public surface:
    check_status() -> dict {is_open, app, workbook, error}
    run_macro(name) -> dict {ok, output, error}
"""

from __future__ import annotations

import platform
import shlex
import subprocess
from typing import Optional


IS_MAC = platform.system() == "Darwin"
IS_WIN = platform.system() == "Windows"


# ─────────────────────────────────────────────────────────────────────
# macOS — AppleScript
# ─────────────────────────────────────────────────────────────────────
def _osascript(script: str, timeout: int = 25) -> tuple[bool, str]:
    """Run an AppleScript snippet. Returns (ok, stdout-or-stderr)."""
    try:
        proc = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        if proc.returncode == 0:
            return True, proc.stdout.strip()
        return False, (proc.stderr or proc.stdout).strip()
    except subprocess.TimeoutExpired:
        return False, "AppleScript timed out (25s). Excel may be unresponsive."
    except FileNotFoundError:
        return False, "osascript not found — is this actually macOS?"
    except Exception as exc:                                      # noqa: BLE001
        return False, f"AppleScript error: {exc}"


def _mac_status() -> dict:
    # First check whether Excel is even running, so we don't *launch* it just by asking.
    running_ok, running_out = _osascript(
        'tell application "System Events" to (name of processes) contains "Microsoft Excel"'
    )
    if not running_ok:
        return {"is_open": False, "app": None, "workbook": None,
                "error": running_out, "running": False}
    if running_out.lower() != "true":
        return {"is_open": False, "app": "Microsoft Excel", "workbook": None,
                "error": None, "running": False}

    # Excel is running — ask which workbook is active
    wb_ok, wb_out = _osascript(
        'tell application "Microsoft Excel" to try\n'
        '    return name of active workbook\n'
        'on error\n'
        '    return ""\n'
        'end try'
    )
    if not wb_ok:
        return {"is_open": True, "app": "Microsoft Excel", "workbook": None,
                "error": wb_out, "running": True}
    return {"is_open": bool(wb_out), "app": "Microsoft Excel",
            "workbook": wb_out or None, "error": None, "running": True}


def _mac_run_macro(name: str) -> dict:
    safe = name.replace('"', '')
    # Excel for Mac's command takes the macro name as a direct parameter
    # (per Excel.sdef) — a labelled "macro name" parameter fails to compile.
    script = (
        'tell application "Microsoft Excel"\n'
        '    with timeout of 1800 seconds\n'
        '        try\n'
        f'            run VB Macro "{safe}"\n'
        '            return "ok"\n'
        '        on error errMsg\n'
        '            return "ERR: " & errMsg\n'
        '        end try\n'
        '    end timeout\n'
        'end tell'
    )
    ok, out = _osascript(script, timeout=1800)
    if not ok:
        return {"ok": False, "output": "", "error": out}
    if out.startswith("ERR:"):
        return {"ok": False, "output": "", "error": out[4:].strip()}
    return {"ok": True, "output": out, "error": None}


# ─────────────────────────────────────────────────────────────────────
# Windows — pywin32 (lazy import)
# ─────────────────────────────────────────────────────────────────────
def _win_status() -> dict:
    try:
        import win32com.client                                    # type: ignore
        import pythoncom                                          # type: ignore
        pythoncom.CoInitialize()
        try:
            xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return {"is_open": False, "app": None, "workbook": None,
                    "error": None, "running": False}
        try:
            wb = xl.ActiveWorkbook.Name if xl.ActiveWorkbook else None
        except Exception:
            wb = None
        return {"is_open": bool(wb), "app": "Microsoft Excel",
                "workbook": wb, "error": None, "running": True}
    except ImportError:
        return {"is_open": False, "app": None, "workbook": None,
                "error": "pywin32 not installed — `pip install pywin32`",
                "running": False}


def _win_run_macro(name: str) -> dict:
    try:
        import win32com.client                                    # type: ignore
        import pythoncom                                          # type: ignore
        pythoncom.CoInitialize()
        xl = win32com.client.GetActiveObject("Excel.Application")
        xl.Application.Run(name)
        return {"ok": True, "output": f"Macro `{name}` executed.", "error": None}
    except Exception as exc:                                      # noqa: BLE001
        return {"ok": False, "output": "", "error": str(exc)}


# ─────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────
def check_status() -> dict:
    if IS_MAC:
        return _mac_status()
    if IS_WIN:
        return _win_status()
    return {"is_open": False, "app": None, "workbook": None,
            "error": f"Unsupported OS: {platform.system()}", "running": False}


def run_macro(name: str) -> dict:
    if not name or not name.strip():
        return {"ok": False, "output": "", "error": "Macro name is empty."}
    if IS_MAC:
        return _mac_run_macro(name.strip())
    if IS_WIN:
        return _win_run_macro(name.strip())
    return {"ok": False, "output": "",
            "error": f"Unsupported OS: {platform.system()}"}
