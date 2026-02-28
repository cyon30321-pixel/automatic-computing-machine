"""
데이터 입력 탭 — 지문+문제 붙여넣기, 검수, 저장
"""

import re
import datetime
import tkinter as tk
from tkinter import messagebox, ttk

from exam_bank.constants import SCHOOL_LEVELS, GRADES, EXAM_TYPES, EXAM_YEARS, EXAM_MONTHS
from exam_bank.models.passage import create_passage, find_duplicate
from exam_bank.services.parser import parse_exam_text


class InputTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self._build()

    def _build(self):
        main = tk.Frame(self.frame)
        main.pack(fill="both", expand=True, padx=20, pady=15)
        tag_frame = tk.LabelFrame(main, text=" 분류 태그 ", font=("맑은 고딕", 10, "bold"), pady=8)
        tag_frame.pack(fill="x", pady=(0, 8))
        r1 = tk.Frame(tag_frame); r1.pack(fill="x", padx=10)

        def _combo(parent, label, values, width=7):
            tk.Label(parent, text=label).pack(side="left", padx=(8, 0))
            cb = ttk.Combobox(parent, values=values, width=width, state="readonly")
            cb.pack(side="left", padx=4)
            return cb

        self.combo_cat1 = _combo(r1, "학교급:", SCHOOL_LEVELS, 6)
        self.combo_year = _combo(r1, "학년:", GRADES, 6)
        self.combo_cat2 = _combo(r1, "유형:", EXAM_TYPES, 8)
        self.combo_exam_year = _combo(r1, "년도:", EXAM_YEARS, 7)
        self.combo_exam_month = _combo(r1, "월:", EXAM_MONTHS, 5)

        r2 = tk.Frame(tag_frame); r2.pack(fill="x", padx=10, pady=8)
        tk.Label(r2, text="학교명/출판사:").pack(side="left")
        self.ent_pub = tk.Entry(r2, width=20); self.ent_pub.pack(side="left", padx=5)
        tk.Label(r2, text="추가 태그:").pack(side="left", padx=(10, 0))
        self.ent_extra_tags = tk.Entry(r2, width=30); self.ent_extra_tags.pack(side="left", padx=5)

        tk.Label(main, text="지문+문제 붙여넣기  (지문과 문제 사이 빈 줄 필수)",
                 font=("맑은 고딕", 10, "bold"), fg="#2563eb").pack(anchor="w", pady=(5, 0))
        self.txt_input = tk.Text(main, font=("Consolas", 10), height=14)
        self.txt_input.pack(fill="both", expand=True, pady=4)

        tk.Label(main, text="정답 및 해설 (선택 — 별도 정답지로 출력)",
                 font=("맑은 고딕", 10, "bold"), fg="#16a34a").pack(anchor="w", pady=(4, 0))
        self.txt_answer = tk.Text(main, font=("Consolas", 10), height=5)
        self.txt_answer.pack(fill="x", pady=4)

        tk.Button(main, text="지문 자동 분해 및 검수", bg="#eab308", fg="#1e293b",
                  font=("맑은 고딕", 12, "bold"), height=2, command=self._open_preview).pack(fill="x")

    def _open_preview(self):
        raw = self.txt_input.get("1.0", tk.END).strip()
        ans_raw = self.txt_answer.get("1.0", tk.END).strip()
        if not raw: return messagebox.showwarning("경고", "텍스트를 입력하세요.")
        parsed = parse_exam_text(raw)
        if not parsed: return messagebox.showwarning("오류", "유효한 지문을 찾지 못했습니다.")

        dup_count = 0
        for item in parsed:
            dup = find_duplicate(self.cfg, item["passage"])
            if dup: item["_duplicate"] = dup; dup_count += 1

        top = tk.Toplevel(self.app.root)
        top.title(f"검수 — {len(parsed)}개 지문 발견" + (f" (중복 {dup_count}개)" if dup_count else ""))
        top.geometry("920x720"); top.attributes("-topmost", True)

        def do_save():
            saved = 0
            for item in parsed:
                if item.get("_duplicate"): continue
                extra = f"{item['tag']} {self.ent_extra_tags.get().strip()}".strip() \
                    if item["tag"] else self.ent_extra_tags.get().strip()
                create_passage(self.cfg, content=item["passage"],
                    category1=self.combo_cat1.get(), category2=self.combo_cat2.get(),
                    school_year=self.combo_year.get(), exam_year=self.combo_exam_year.get(),
                    exam_month=self.combo_exam_month.get(), publisher=self.ent_pub.get(),
                    extra_tags=extra, answer_text=ans_raw, questions=item["questions"])
                saved += 1
            msg = f"{saved}개 지문 저장 완료!"
            if dup_count: msg += f"\n({dup_count}개 중복 건너뛰)"
            messagebox.showinfo("완료", msg); top.destroy()
            self.txt_input.delete("1.0", tk.END); self.txt_answer.delete("1.0", tk.END)
            self.app.refresh_all()

        # 버튼을 먼저 pack (항상 하단에 보임)
        tk.Button(top, text="✅ 최종 저장", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 12, "bold"), height=2, command=do_save).pack(fill="x", padx=16, pady=8, side="bottom")

        txt = tk.Text(top, font=("Consolas", 10), bg="#f8fafc")
        txt.pack(fill="both", expand=True, padx=16, pady=8)
        for i, item in enumerate(parsed, 1):
            dup_warn = " ⚠️ 중복!" if item.get("_duplicate") else ""
            txt.insert(tk.END, f"{'='*20} [지문 {i}]{dup_warn} {'='*20}\n\n")
            txt.insert(tk.END, item["passage"] + "\n\n")
            for q in item["questions"]:
                num = f"{q['q_num']}. " if q["q_num"] != "-" else ""
                txt.insert(tk.END, f"{num}{q['content']}\n\n")
            txt.insert(tk.END, "\n")
