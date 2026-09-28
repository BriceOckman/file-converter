#!/usr/bin/env python3
"""One-step launcher for File Convertor.

Just run:  python3 start.py

It creates a virtual environment on first run, installs the requirements
once (and again only if requirements.txt changes), then opens the app.
You never have to touch a venv or pip yourself.
"""

import hashlib
import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENV = ROOT / ".venv"
REQS = ROOT / "requirements.txt"
MARKER = VENV / ".reqs.sha256"


def venv_python() -> Path:
    if os.name == "nt":
        return VENV / "Scripts" / "python.exe"
    return VENV / "bin" / "python"


def ensure_venv() -> Path:
    py = venv_python()
    if not py.exists():
        print("First run: creating a virtual environment…", flush=True)
        subprocess.run([sys.executable, "-m", "venv", str(VENV)], check=True)
    return py


def ensure_requirements(py: Path) -> None:
    want = hashlib.sha256(REQS.read_bytes()).hexdigest()
    have = MARKER.read_text().strip() if MARKER.exists() else ""
    if have == want:
        return  # already installed, skip the slow pip check
    print("First run: installing requirements (one-time download)…", flush=True)
    subprocess.run(
        [str(py), "-m", "pip", "install", "--quiet",
         "--disable-pip-version-check", "-r", str(REQS)],
        check=True,
    )
    MARKER.write_text(want)
    print("Setup done.", flush=True)


def check_tkinter(py: Path) -> None:
    r = subprocess.run([str(py), "-c", "import tkinter"],
                       capture_output=True)
    if r.returncode != 0:
        print("ERROR: this Python is missing tkinter (needed for the GUI).")
        if sys.platform == "darwin":
            print("Fix: install Python from python.org (it bundles tkinter).")
        elif sys.platform.startswith("linux"):
            print("Fix: sudo apt install python3-tk   (or your distro's equivalent)")
        else:
            print("Fix: reinstall Python with the 'tcl/tk and IDLE' option enabled.")
        sys.exit(1)


def main() -> None:
    py = ensure_venv()
    ensure_requirements(py)
    check_tkinter(py)
    print("Starting File Convertor…", flush=True)
    sys.stdout.flush()
    sys.stderr.flush()
    os.execv(str(py), [str(py), str(ROOT / "app.py")])


if __name__ == "__main__":
    main()
