"""
공용 위젯 — 툴팁, 상태바
"""

import textwrap
import tkinter as tk


class ToolTip:
    """트리뷰 마우스오버 툴팁."""
    def __init__(self, root):
        self.root = root
        self.win = None
        self.label = None

    def show(self, text, x, y):
        wrapped = textwrap.fill(text[:500] + ("..." if len(text) > 500 else ""), width=60)
        if self.win:
            self.label.config(text=wrapped)
            self.win.geometry(f"+{x+15}+{y+15}")
        else:
            self.win = tk.Toplevel(self.root)
            self.win.wm_overrideredirect(True)
            self.win.geometry(f"+{x+15}+{y+15}")
            self.label = tk.Label(
                self.win, text=wrapped, justify="left",
                bg="#fefce8", relief="solid", borderwidth=1, font=("맑은 고딕", 9),
            )
            self.label.pack(ipadx=5, ipady=5)

    def hide(self):
        if self.win:
            self.win.destroy()
            self.win = None
            self.label = None


class StatusBar(tk.Frame):
    """하단 상태바."""
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg="#e2e8f0", height=26, **kwargs)
        self.pack_propagate(False)
        self.label = tk.Label(
            self, text="", font=("맑은 고딕", 8),
            bg="#e2e8f0", fg="#475569",
        )
        self.label.pack(side="left", padx=12)

    def set_text(self, text):
        self.label.config(text=text)
