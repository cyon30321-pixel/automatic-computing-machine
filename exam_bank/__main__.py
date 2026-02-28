"""
패키지 실행 진입점
python -m exam_bank
"""

import sys
import os
import tkinter as tk

# 패키지 디렉터리를 PATH에 추가
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from exam_bank.app import ExamBankApp


def main():
    root = tk.Tk()
    app = ExamBankApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
