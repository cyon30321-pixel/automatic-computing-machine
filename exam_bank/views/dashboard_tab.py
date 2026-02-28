"""
대시보드 탭 — 앱 시작 시 전체 현황 요약
"""

import tkinter as tk
from tkinter import ttk

from exam_bank.models.database import db_stats, db_conn
from exam_bank.models.student import list_students, list_exam_records


class DashboardTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self._build()

    def _build(self):
        main = tk.Frame(self.frame)
        main.pack(fill="both", expand=True, padx=30, pady=20)

        # 제목
        tk.Label(
            main, text="📊 대시보드", font=("맑은 고딕", 16, "bold"), fg="#1e293b"
        ).pack(anchor="w", pady=(0, 15))

        # 통계 카드
        card_frame = tk.Frame(main)
        card_frame.pack(fill="x", pady=(0, 20))

        self.cards = {}
        for i, (key, label, color) in enumerate([
            ("passages", "지문", "#2563eb"),
            ("questions", "문항", "#16a34a"),
            ("students", "학생", "#ea580c"),
            ("exams", "시험 기록", "#7c3aed"),
        ]):
            card = tk.Frame(card_frame, bg=color, padx=20, pady=15)
            card.grid(row=0, column=i, padx=8, sticky="nsew")
            card_frame.columnconfigure(i, weight=1)

            tk.Label(card, text=label, font=("맑은 고딕", 10), bg=color, fg="white").pack()
            lbl = tk.Label(card, text="0", font=("맑은 고딕", 24, "bold"), bg=color, fg="white")
            lbl.pack()
            self.cards[key] = lbl

        # 최근 활동
        act_frame = tk.LabelFrame(main, text=" 최근 시험 기록 ", font=("맑은 고딕", 10, "bold"))
        act_frame.pack(fill="both", expand=True, pady=8)

        cols = ("날짜", "학생", "제목", "문항수", "채점", "점수")
        self.tree_recent = ttk.Treeview(act_frame, columns=cols, show="headings", height=8)
        for col, w in zip(cols, [90, 80, 200, 60, 60, 60]):
            self.tree_recent.heading(col, text=col)
            self.tree_recent.column(col, width=w)
        self.tree_recent.pack(fill="both", expand=True, padx=8, pady=8)

        # 새로고침 버튼
        tk.Button(
            main, text="새로고침", command=self.refresh,
            bg="#64748b", fg="white", font=("맑은 고딕", 9),
        ).pack(anchor="e", pady=4)

    def refresh(self):
        """대시보드 데이터 갱신."""
        try:
            stats = db_stats(self.cfg)
            for key, lbl in self.cards.items():
                lbl.config(text=str(stats.get(key, 0)))
        except Exception:
            pass

        # 최근 시험 기록
        for item in self.tree_recent.get_children():
            self.tree_recent.delete(item)

        try:
            records = list_exam_records(self.cfg)[:20]
            students = {s["id"]: s["name"] for s in list_students(self.cfg)}
            for r in records:
                graded = "✅" if r["is_graded"] else "—"
                score = f"{r['score']}%" if r["is_graded"] else "—"
                self.tree_recent.insert("", "end", values=(
                    r["exam_date"],
                    students.get(r["student_id"], "?"),
                    r.get("exam_title", ""),
                    r["total_questions"],
                    graded,
                    score,
                ))
        except Exception:
            pass
