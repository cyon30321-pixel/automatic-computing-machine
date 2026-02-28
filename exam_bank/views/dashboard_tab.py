"""
대시보드 탭 — 홈 통계, 최근 활동
"""

import tkinter as tk
from tkinter import ttk

from exam_bank.models.database import db_conn, db_stats


class DashboardTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self._build()

    def _build(self):
        # 통계 카드
        cf = tk.Frame(self.frame)
        cf.pack(fill="x", padx=20, pady=20)

        self.cards = {}
        colors = [
            ("passages", "지문", "#3b82f6"),
            ("questions", "문제", "#22c55e"),
            ("students", "학생", "#f97316"),
            ("exams", "시험", "#a855f7"),
        ]
        for key, label, color in colors:
            card = tk.Frame(cf, bg=color, width=200, height=100)
            card.pack(side="left", padx=12, pady=8)
            card.pack_propagate(False)
            tk.Label(card, text=label, bg=color, fg="white", font=("맑은 고딕", 11)).pack(pady=(12, 0))
            num_lbl = tk.Label(card, text="0", bg=color, fg="white", font=("맑은 고딕", 24, "bold"))
            num_lbl.pack()
            self.cards[key] = num_lbl

        # 최근 활동
        tk.Label(self.frame, text="최근 시험 기록", font=("맑은 고딕", 11, "bold")).pack(anchor="w", padx=20, pady=(16, 4))

        cols = ("날짜", "학생", "시험명", "문항수", "채점", "점수")
        self.tree = ttk.Treeview(self.frame, columns=cols, show="headings", height=8)
        for c, w in zip(cols, [100, 80, 180, 60, 50, 60]):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w)
        self.tree.pack(fill="x", padx=20, pady=4)

        tk.Button(self.frame, text="새로고침", command=self.refresh).pack(pady=8)

    def refresh(self):
        try:
            s = db_stats(self.cfg)
            for key, lbl in self.cards.items():
                lbl.config(text=str(s.get(key, 0)))
        except Exception:
            pass

        for i in self.tree.get_children():
            self.tree.delete(i)

        try:
            with db_conn(self.cfg) as conn:
                cur = conn.cursor()
                rows = cur.execute(
                    """SELECT er.exam_date, s.name, er.exam_title,
                              er.total_questions, er.is_graded, er.score
                       FROM exam_records er
                       JOIN students s ON er.student_id = s.id
                       ORDER BY er.id DESC LIMIT 20"""
                ).fetchall()
                for r in rows:
                    graded = "✅" if r[4] else "⬜"
                    score = f"{r[5]}%" if r[4] else "-"
                    self.tree.insert("", "end", values=(r[0], r[1], r[2], r[3], graded, score))
        except Exception:
            pass
