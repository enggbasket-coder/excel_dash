#!/usr/bin/env bash
# Launch the redesigned Excel VBA Dashboard.
# Runs on port 8502 so it doesn't clash with your existing app on 8501.
set -e
cd "$(dirname "$0")"

# Optional: create / activate a venv on first run
if [ ! -d ".venv" ]; then
    python3 -m venv .venv
    ./.venv/bin/pip install --upgrade pip
    ./.venv/bin/pip install -r requirements.txt
fi

./.venv/bin/streamlit run app.py --server.port 8502
