"""
공용 위젯 — 툴팁, 상태바
"""

import tkinter as tk


class ToolTip:
    def __init__(self, root):
        self.tw = None
        self.root = root

    def show(self, text, x, y):
        self.hide()
        self.tw = tk.Toplevel(self.root)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry(f"+{x+15}+{y+10}")
        lbl = tk.Label(
            self.tw, text=text[:300], justify="left",
            bg="#fffbe6", fg="#333", relief="solid", bd=1,
            font=("Consolas", 9), wraplength=400,
        )
        lbl.pack()

    def hide(self):
        if self.tw:
            self.tw.destroy()
            self.tw = None


class StatusBar(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bd=1, relief="sunken")
        self.lbl = tk.Label(
            self, text="준비 완료", anchor="w",
            font=("Consolas", 9), fg="#475569",
        )
        self.lbl.pack(fill="x", padx=6, pady=2)

    def set_text(self, text):
        self.lbl.config(text=text)
