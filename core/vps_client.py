"""
VPS client — talks to the Excel Audit API proxy at excel.epcmode.in.

All LLM calls from the Streamlit app go through here.
The Google API key lives on the VPS only — never on the customer's machine.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

import requests

# ─────────────────────────────────────────────────────────────
VPS_BASE_URL  = "https://excel.epcmode.in"
LLM_ENDPOINT  = f"{VPS_BASE_URL}/api/llm"
VAL_ENDPOINT  = f"{VPS_BASE_URL}/api/license/validate"
TIMEOUT       = 120          # seconds — LLM calls can be slow
LICENSE_FILE  = Path.home() / ".excel_audit_license"


# ─────────────────────────────────────────────────────────────
#  License key persistence
# ─────────────────────────────────────────────────────────────
def load_license_key() -> str:
    """Read saved license key from ~/.excel_audit_license"""
    if LICENSE_FILE.exists():
        return LICENSE_FILE.read_text(encoding="utf-8").strip()
    return ""


def save_license_key(key: str) -> None:
    """Persist license key to disk."""
    LICENSE_FILE.write_text(key.strip(), encoding="utf-8")
    # Restrict permissions on Mac/Linux (Windows ignores chmod)
    try:
        LICENSE_FILE.chmod(0o600)
    except Exception:
        pass


def clear_license_key() -> None:
    if LICENSE_FILE.exists():
        LICENSE_FILE.unlink()


# ─────────────────────────────────────────────────────────────
#  Validate a key against the VPS
# ─────────────────────────────────────────────────────────────
def validate_license(key: str) -> dict:
    """
    Returns dict:
      { valid: bool, customer_name, plan, quota_monthly,
        calls_this_month, expires_at, message }
    Never raises — returns {valid: False, message: <error>} on failure.
    """
    if not key:
        return {"valid": False, "message": "No license key provided."}
    try:
        resp = requests.get(
            VAL_ENDPOINT,
            headers={"X-License-Key": key},
            timeout=15,
        )
        if resp.status_code == 200:
            return resp.json()
        return {"valid": False, "message": f"Server returned {resp.status_code}: {resp.text[:200]}"}
    except requests.ConnectionError:
        return {"valid": False, "message": "Cannot reach excel.epcmode.in — check your internet connection."}
    except requests.Timeout:
        return {"valid": False, "message": "Server timed out during validation."}
    except Exception as exc:
        return {"valid": False, "message": str(exc)}


# ─────────────────────────────────────────────────────────────
#  Main LLM call
# ─────────────────────────────────────────────────────────────
def ask_llm(
    system_prompt: str,
    user_prompt: str,
    action: str = "general",
    license_key: Optional[str] = None,
) -> dict:
    """
    Send a prompt to the VPS LLM proxy.

    Returns:
        { ok: bool, text: str, calls_remaining: int, error: str }
    """
    key = license_key or load_license_key()
    if not key:
        return {
            "ok": False,
            "text": "",
            "calls_remaining": 0,
            "error": "No license key. Enter your key in the LLM Audit tab.",
        }

    payload = {
        "system_prompt": system_prompt,
        "user_prompt":   user_prompt,
        "action":        action,
    }

    try:
        resp = requests.post(
            LLM_ENDPOINT,
            json=payload,
            headers={
                "X-License-Key": key,
                "Content-Type":  "application/json",
            },
            timeout=TIMEOUT,
        )
    except requests.ConnectionError:
        return {
            "ok": False, "text": "", "calls_remaining": 0,
            "error": "Cannot reach excel.epcmode.in — check your internet.",
        }
    except requests.Timeout:
        return {
            "ok": False, "text": "", "calls_remaining": 0,
            "error": f"Request timed out after {TIMEOUT}s.",
        }
    except Exception as exc:
        return {"ok": False, "text": "", "calls_remaining": 0, "error": str(exc)}

    if resp.status_code == 200:
        data = resp.json()
        return {
            "ok": True,
            "text": data.get("text", ""),
            "calls_remaining": data.get("calls_remaining", 0),
            "error": "",
        }

    # Map HTTP errors to readable messages
    msg_map = {
        401: "License key invalid, expired, or quota exceeded.",
        403: "Access forbidden.",
        429: "Too many requests — slow down.",
        502: "AI service temporarily unavailable.",
        503: "Server not configured — contact admin.",
        504: "AI service timed out.",
    }
    msg = msg_map.get(resp.status_code, f"Server error {resp.status_code}: {resp.text[:200]}")
    return {"ok": False, "text": "", "calls_remaining": 0, "error": msg}
