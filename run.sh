#!/bin/bash
# Launch the File Convertor app.
cd "$(dirname "$0")"
if [ -d ".venv" ]; then
    .venv/bin/python app.py
else
    python3 app.py
fi
