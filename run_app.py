#!/usr/bin/env python3
"""
Launcher for the Uganda Report Card Web App.
Run:  python run_app.py
Then open:  http://127.0.0.1:5000
"""

import sys
import os
from pathlib import Path

# Make Flask (user install) and the project root importable
sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, "/root/.local/lib/python3.12/site-packages")

from app.app import app

if __name__ == "__main__":
    print("=" * 60)
    print("  🇺🇬  UGANDA SECONDARY SCHOOL REPORT CARD SYSTEM")
    print("  Web Application")
    print("=" * 60)
    print()
    print("  Open in your browser:")
    print("    →  http://127.0.0.1:5000")
    print("    →  http://localhost:5000")
    print()
    print("  Press Ctrl+C to stop the server.")
    print("=" * 60)
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
