"""
메인 앱 v5.0 — 루트 윈도우, 7개 탭 라우팅, 상태바
"""

import tkinter as tk
from tkinter import messagebox, ttk

from exam_bank.constants import APP_NAME
from exam_bank.config import load_config, save_config, setup_korean_font
from exam_bank.models.database import init_db, db_stats
from exam_bank.views.widgets import StatusBar
from exam_bank.views.dashboard_tab import DashboardTab
from exam_bank.views.input_tab import InputTab
from exam_bank.views.bank_tab import BankTab
from exam_bank.views.exam_tab import ExamTab
from exam_bank.views.ai_prompt_tab import AIPromptTab
from exam_bank.views.analysis_tab import AnalysisTab
from exam_bank.views.config_tab import ConfigTab


class ExamBankApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("1400x1020")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        setup_korean_font()
        self.cfg = load_config()
        init_db(self.cfg)
        self._build_ui()
        self._bind_shortcuts()
        self.refresh_all()

    def _build_ui(self):
        self.status_bar = StatusBar(self.root)
        self.status_bar.pack(fill="x", side="bottom")

        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=8)

        # 7개 탭 생성
        self.dashboard_tab = DashboardTab(self.tabs, self.cfg, self)
        self.input_tab = InputTab(self.tabs, self.cfg, self)
        self.bank_tab = BankTab(self.tabs, self.cfg, self)
        self.exam_tab = ExamTab(self.tabs, self.cfg, self)
        self.ai_prompt_tab = AIPromptTab(self.tabs, self.cfg, self)
        self.analysis_tab = AnalysisTab(self.tabs, self.cfg, self)
        self.config_tab = ConfigTab(self.tabs, self.cfg, self)

        self.tabs.add(self.dashboard_tab.frame, text="  📊 대시보드  ")
        self.tabs.add(self.input_tab.frame, text="  📝 데이터 입력  ")
        self.tabs.add(self.bank_tab.frame, text="  📚 문제 관리  ")
        self.tabs.add(self.exam_tab.frame, text="  🖨️ 시험지 생성  ")
        self.tabs.add(self.ai_prompt_tab.frame, text="  🤖 AI 프롬프트  ")
        self.tabs.add(self.analysis_tab.frame, text="  🔍 오답 분석  ")
        self.tabs.add(self.config_tab.frame, text="  ⚙️ 설정  ")

    def _bind_shortcuts(self):
        self.root.bind("<Control-f>", lambda e: (
            self.tabs.select(self.bank_tab.frame),
            self.bank_tab.ent_search.focus_set()
        ))
        self.root.bind("<Control-s>", lambda e: self.bank_tab._generate("exam"))

    def _on_close(self):
        if self.bank_tab.cart:
            if not messagebox.askyesno("종료 확인", "장바구니에 항목이 있습니다.\n정말 종료하시겠습니까?"):
                return
        save_config(self.cfg)
        self.root.destroy()

    def refresh_status(self):
        try:
            s = db_stats(self.cfg)
            cart_count = sum(1 if not v else len(v) for v in self.bank_tab.cart.values())
            self.status_bar.set_text(
                f"DB: 지문 {s['passages']}개 | 문제 {s['questions']}개 | "
                f"학생 {s['students']}명 | 시험 {s['exams']}건    "
                f"장바구니: {cart_count}건"
            )
        except Exception:
            self.status_bar.set_text("DB 연결 확인 필요")

    def refresh_all(self):
        self.refresh_status()
        try:
            self.dashboard_tab.refresh()
        except Exception:
            pass
        try:
            self.bank_tab.refresh()
        except Exception:
            pass
        try:
            self.config_tab.refresh_stats()
        except Exception:
            pass
