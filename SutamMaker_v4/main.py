"""
SutamMaker v4 — main.py
━━━━━━━━━━━━━━━━━━━━━━━
Application entry-point.

Sets up sys.path so that every module can use simple imports:
    from config import …
    from core.parser_engine import …
    from data.data_manager import …
    import exam_db            # lives in the repo root (parent dir)
"""

import sys
import os

# ── path setup ──────────────────────────────────────────────
# 1) This directory (SutamMaker_v4/) → allows "from config import …"
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

# 2) Parent directory (repo root) → allows "import exam_db"
_PARENT_DIR = os.path.dirname(_THIS_DIR)
if _PARENT_DIR not in sys.path:
    sys.path.insert(1, _PARENT_DIR)

# ── imports (after path setup) ──────────────────────────────
import tkinter as tk  # noqa: E402

try:
    from tkinterdnd2 import TkinterDnD
except ImportError:
    TkinterDnD = None

from ui.ui_main import SutamMakerApp  # noqa: E402


def main():
    if TkinterDnD:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("Info: tkinterdnd2 미설치 — Drag & Drop 기능이 비활성화됩니다.")

    SutamMakerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
