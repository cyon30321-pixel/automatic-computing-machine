"""
엔트리포인트: python -m exam_bank
"""

import os
import sys
import tkinter as tk

# 패키지 경로 설정
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

os.chdir(SCRIPT_DIR)

from exam_bank.app import ExamBankApp

if __name__ == "__main__":
    root = tk.Tk()
    app = ExamBankApp(root)
    root.mainloop()
