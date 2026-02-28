"""
AI 오답 분석 탭 — 답안 기입/채점, 자동 취약점 분석, AI 프롬프트, 시각화
"""

import re
import json
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog

from exam_bank.models.database import db_conn
from exam_bank.models.student import (
    create_student, update_student, delete_student, list_students, get_student,
    list_exam_records, get_exam_record, submit_answers,
    get_student_wrong_answers, get_student_weakness_summary,
)
from exam_bank.models.analysis import (
    save_analysis, save_auto_analysis, get_analysis_history, get_latest_weakness,
)
from exam_bank.models.passage import search_passages
from exam_bank.services.analytics import (
    chart_progress, chart_weakness, export_csv, generate_ai_prompt,
)
from exam_bank.config import open_directory


class AnalysisTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self.current_student_id = None
        self.current_student_name = ""
        self._build()

    def _build(self):
        self.frame.columnconfigure(0, weight=1)
        self.frame.columnconfigure(1, weight=3)
        self.frame.columnconfigure(2, weight=2)

        # ── 좌: 학생 선택 ──
        f_stu = tk.LabelFrame(self.frame, text=" 1. 학생 선택 ", font=("맑은 고딕", 10, "bold"))
        f_stu.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)

        self.ent_student = tk.Entry(f_stu)
        self.ent_student.pack(fill="x", padx=6, pady=4)
        self.ent_student.bind("<Return>", lambda e: self._add_student())

        btn_stu = tk.Frame(f_stu); btn_stu.pack(fill="x", padx=6, pady=2)
        tk.Button(btn_stu, text="추가", command=self._add_student, width=6).pack(side="left", padx=2)
        tk.Button(btn_stu, text="삭제", command=self._delete_student, bg="#dc2626", fg="white", width=6).pack(side="left", padx=2)
        tk.Button(btn_stu, text="수정", command=self._rename_student, width=6).pack(side="left", padx=2)

        self.list_students = tk.Listbox(f_stu)
        self.list_students.pack(fill="both", expand=True, padx=6, pady=6)
        self.list_students.bind("<<ListboxSelect>>", self._on_student_select)

        # ── 중: 채점 & 분석 ──
        f_mid = tk.LabelFrame(self.frame, text=" 2. 채점 & 분석 ", font=("맑은 고딕", 10, "bold"))
        f_mid.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)

        # 시험 기록 선택
        exam_row = tk.Frame(f_mid); exam_row.pack(fill="x", padx=10, pady=4)
        tk.Label(exam_row, text="시험 기록:").pack(side="left")
        self.combo_exam = ttk.Combobox(exam_row, values=[], width=35, state="readonly")
        self.combo_exam.pack(side="left", padx=4)
        tk.Button(exam_row, text="불러오기", command=self._load_exam).pack(side="left", padx=4)

        # 답안 입력 영역
        tk.Label(f_mid, text="답안 입력 (문항별 정답 입력 후 채점)", font=("맑은 고딕", 9, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
        self.answer_frame = tk.Frame(f_mid)
        self.answer_frame.pack(fill="both", expand=True, padx=10, pady=4)

        # 스크롤
        canvas = tk.Canvas(self.answer_frame, height=200)
        scrollbar = tk.Scrollbar(self.answer_frame, orient="vertical", command=canvas.yview)
        self.answer_inner = tk.Frame(canvas)
        self.answer_inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.answer_inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.answer_entries = {}  # {question_id: entry_widget}
        self.current_exam_id = None

        btn_row = tk.Frame(f_mid); btn_row.pack(fill="x", padx=10, pady=4)
        tk.Button(btn_row, text="채점하기", bg="#ea580c", fg="white", font=("맑은 고딕", 11, "bold"), height=2, command=self._grade_exam).pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(btn_row, text="자동 취약점 분석", bg="#7c3aed", fg="white", font=("맑은 고딕", 11, "bold"), height=2, command=self._auto_analyze).pack(side="left", fill="x", expand=True, padx=2)

        # 수동 분석 (레거시)
        manual_frame = tk.LabelFrame(f_mid, text=" 수동 분석 입력 ", font=("맑은 고딕", 9))
        manual_frame.pack(fill="x", padx=10, pady=4)
        mr = tk.Frame(manual_frame); mr.pack(fill="x", padx=6, pady=4)
        tk.Label(mr, text="취약점:").pack(side="left")
        self.ent_weakness = tk.Entry(mr, width=12, fg="#dc2626"); self.ent_weakness.pack(side="left", padx=4)
        tk.Label(mr, text="점수:").pack(side="left")
        self.ent_score = tk.Entry(mr, width=6); self.ent_score.pack(side="left", padx=4)
        self.txt_feedback = tk.Text(manual_frame, height=4, bg="#f0f9ff", font=("맑은 고딕", 9))
        self.txt_feedback.pack(fill="x", padx=6, pady=2)
        tk.Button(manual_frame, text="수동 저장", command=self._save_manual, bg="#64748b", fg="white").pack(fill="x", padx=6, pady=4)

        # ── 우: 시각화 & AI ──
        f_viz = tk.LabelFrame(self.frame, text=" 3. 시각화 & AI ", font=("맑은 고딕", 10, "bold"))
        f_viz.grid(row=0, column=2, sticky="nsew", padx=8, pady=8)

        tk.Button(f_viz, text="성취도 추이 (Line)", bg="#0891b2", fg="white", font=("맑은 고딕", 10, "bold"), height=2, command=self._chart_progress).pack(fill="x", padx=16, pady=4)
        tk.Button(f_viz, text="취약점 분석 (Bar)", bg="#7c3aed", fg="white", font=("맑은 고딕", 10, "bold"), height=2, command=self._chart_weakness).pack(fill="x", padx=16, pady=4)
        tk.Button(f_viz, text="CSV 내보내기", bg="#334155", fg="white", font=("맑은 고딕", 10, "bold"), height=2, command=self._export_csv).pack(fill="x", padx=16, pady=4)

        tk.Label(f_viz, text="").pack(pady=4)  # spacer

        tk.Button(
            f_viz, text="🤖 AI 맞춤 프롬프트 생성", bg="#ea580c", fg="white",
            font=("맑은 고딕", 11, "bold"), height=3, command=self._generate_ai_prompt,
        ).pack(fill="x", padx=16, pady=8)

        tk.Label(f_viz, text="타겟 취약점:").pack(pady=(12, 0))
        self.lbl_weakness = tk.Label(f_viz, text="없음", font=("Consolas", 12, "bold"), fg="#dc2626")
        self.lbl_weakness.pack(pady=4)

        tk.Button(
            f_viz, text="취약점 타겟 보충문제 추출", bg="#16a34a", fg="white",
            font=("맑은 고딕", 10, "bold"), height=2, command=self._make_remedial,
        ).pack(fill="x", padx=16, pady=4)

    # ── 학생 CRUD ─────────────────────────────

    def _add_student(self):
        name = self.ent_student.get().strip()
        if not name: return messagebox.showwarning("오류", "이름을 입력하세요.")
        create_student(self.cfg, name, grade="공통")
        self.ent_student.delete(0, tk.END)
        self._refresh_students()
        self.app.refresh_status()

    def _delete_student(self):
        if not self.current_student_id: return messagebox.showwarning("알림", "학생을 선택하세요.")
        if not messagebox.askyesno("확인", f"'{self.current_student_name}' 학생과 모든 기록을 삭제하시겠습니까?"): return
        delete_student(self.cfg, self.current_student_id)
        self.current_student_id = None; self.current_student_name = ""
        self._refresh_students(); self.app.refresh_status()

    def _rename_student(self):
        if not self.current_student_id: return messagebox.showwarning("알림", "학생을 선택하세요.")
        new = simpledialog.askstring("이름 수정", "새 이름:", parent=self.app.root, initialvalue=self.current_student_name)
        if not new or not new.strip(): return
        update_student(self.cfg, self.current_student_id, name=new.strip())
        self.current_student_name = new.strip()
        self._refresh_students()

    def _refresh_students(self):
        self.list_students.delete(0, tk.END)
        for s in list_students(self.cfg):
            self.list_students.insert(tk.END, f"[{s['id']}] {s['name']} ({s.get('grade','')})")

    def _on_student_select(self, event):
        sel = self.list_students.curselection()
        if not sel: return
        val = self.list_students.get(sel[0])
        self.current_student_id = int(re.search(r"\[(\d+)\]", val).group(1))
        self.current_student_name = re.search(r"\]\s*([^(]*)", val).group(1).strip()

        # 시험 기록 콤보 갱신
        records = list_exam_records(self.cfg, self.current_student_id)
        labels = []
        for r in records:
            graded = "✅" if r["is_graded"] else "⬜"
            labels.append(f"[{r['id']}] {r['exam_date']} | {r.get('exam_title','')} | {graded} {r.get('score',0)}%")
        self.combo_exam["values"] = labels
        if labels: self.combo_exam.current(0)

        # 취약점 표시
        weakness = get_latest_weakness(self.cfg, self.current_student_id)
        self.lbl_weakness.config(text=f"Target: {weakness}" if weakness else "없음")

    # ── 답안 기입 & 채점 ─────────────────────

    def _load_exam(self):
        val = self.combo_exam.get()
        m = re.search(r"\[(\d+)\]", val)
        if not m: return messagebox.showwarning("알림", "시험 기록을 선택하세요.")
        exam_id = int(m.group(1))
        self.current_exam_id = exam_id

        record = get_exam_record(self.cfg, exam_id)
        if not record or not record.get("items"):
            return messagebox.showinfo("안내", "시험 문항 데이터가 없습니다.")

        # 기존 위젯 삭제
        for w in self.answer_inner.winfo_children():
            w.destroy()
        self.answer_entries.clear()

        for item in record["items"]:
            qid = item["question_id"]
            row = tk.Frame(self.answer_inner); row.pack(fill="x", pady=2)

            # 문항 번호 + 내용 미리보기
            q_text = f"{item['display_order']}. {(item.get('q_content',''))[:50]}"
            tk.Label(row, text=q_text, width=55, anchor="w", font=("맑은 고딕", 9)).pack(side="left")

            # 정답 표시
            correct = item.get("correct_answer", "")
            tk.Label(row, text=f"정답:{correct}", width=8, fg="#16a34a", font=("맑은 고딕", 8)).pack(side="left")

            # 학생 답안 입력
            ent = tk.Entry(row, width=5, font=("맑은 고딕", 10, "bold"))
            ent.pack(side="left", padx=4)
            if item.get("student_answer"):
                ent.insert(0, item["student_answer"])
            self.answer_entries[qid] = ent

            # 채점 결과
            if item["is_correct"] == 1:
                tk.Label(row, text="✅", fg="#16a34a").pack(side="left")
            elif item["is_correct"] == 0:
                tk.Label(row, text="❌", fg="#dc2626").pack(side="left")

    def _grade_exam(self):
        if not self.current_exam_id: return messagebox.showwarning("알림", "시험 기록을 먼저 불러오세요.")
        if not self.answer_entries: return

        answers = {}
        for qid, ent in self.answer_entries.items():
            val = ent.get().strip()
            if val:
                answers[qid] = val

        if not answers: return messagebox.showwarning("알림", "답안을 1개 이상 입력하세요.")

        result = submit_answers(self.cfg, self.current_exam_id, answers)
        messagebox.showinfo("채점 완료",
            f"총 {result['total']}문항\n"
            f"정답: {result['correct']}개 | 오답: {result['wrong']}개\n"
            f"점수: {result['score']}%")

        # 화면 새로고침
        self._load_exam()
        self._on_student_select(None)

    def _auto_analyze(self):
        if not self.current_student_id: return messagebox.showwarning("알림", "학생을 선택하세요.")
        result = save_auto_analysis(self.cfg, self.current_student_id)
        if not result:
            return messagebox.showinfo("안내", "오답 데이터가 없어 자동 분석이 불가합니다.\n먼저 채점을 해주세요.")
        self.lbl_weakness.config(text=f"Target: {result['tags']}")
        detail = "\n".join(f"- {d['tag']}: 오답률 {d['wrong_rate']}%" for d in result["details"])
        messagebox.showinfo("자동 분석 완료", f"취약점: {result['tags']}\n이해도: {result['score']}점\n\n{detail}")

    def _save_manual(self):
        if not self.current_student_id: return messagebox.showwarning("오류", "학생을 선택하세요.")
        weakness = self.ent_weakness.get().strip()
        score_str = self.ent_score.get().strip()
        feedback = self.txt_feedback.get("1.0", tk.END).strip()
        if not weakness or not score_str.isdigit(): return messagebox.showwarning("오류", "취약점과 점수(숫자)를 입력하세요.")
        save_analysis(self.cfg, self.current_student_id, weakness, int(score_str), feedback, source="manual")
        self.ent_weakness.delete(0, tk.END); self.ent_score.delete(0, tk.END)
        messagebox.showinfo("저장 완료", "수동 분석 저장 완료!")
        self._on_student_select(None)

    # ── 시각화 ────────────────────────────────

    def _chart_progress(self):
        if not self.current_student_id: return messagebox.showwarning("오류", "학생을 선택하세요.")
        ok, msg = chart_progress(self.cfg, self.current_student_id, self.current_student_name)
        if not ok: messagebox.showinfo("안내", msg)

    def _chart_weakness(self):
        if not self.current_student_id: return messagebox.showwarning("오류", "학생을 선택하세요.")
        ok, msg = chart_weakness(self.cfg, self.current_student_id, self.current_student_name)
        if not ok: messagebox.showinfo("안내", msg)

    def _export_csv(self):
        if not self.current_student_id: return messagebox.showwarning("오류", "학생을 선택하세요.")
        path = export_csv(self.cfg, self.current_student_id, self.current_student_name, self.cfg["last_dir"])
        messagebox.showinfo("완료", f"CSV 저장: {path}")
        open_directory(self.cfg["last_dir"])

    # ── AI 프롬프트 ───────────────────────────

    def _generate_ai_prompt(self):
        if not self.current_student_id: return messagebox.showwarning("알림", "학생을 선택하세요.")
        prompt = generate_ai_prompt(self.cfg, self.current_student_id)
        if not prompt: return messagebox.showinfo("안내", "프롬프트 생성에 실패했습니다.")

        # 프롬프트 표시 + 복사
        top = tk.Toplevel(self.app.root); top.title("AI 맞춤 프롬프트"); top.geometry("700x600"); top.attributes("-topmost", True)
        tk.Label(top, text="아래 프롬프트를 복사하여 ChatGPT/Claude에 붙여넣으세요.", font=("맑은 고딕", 10, "bold"), fg="#2563eb").pack(pady=8)
        txt = tk.Text(top, font=("Consolas", 10), bg="#f8fafc"); txt.pack(fill="both", expand=True, padx=16, pady=4)
        txt.insert("1.0", prompt)

        def copy():
            self.app.root.clipboard_clear(); self.app.root.clipboard_append(prompt)
            messagebox.showinfo("복사 완료", "클립보드에 복사되었습니다!")

        tk.Button(top, text="📋 클립보드에 복사", bg="#2563eb", fg="white", font=("맑은 고딕", 11, "bold"), height=2, command=copy).pack(fill="x", padx=16, pady=8)

    # ── 보충문제 추출 ─────────────────────────

    def _make_remedial(self):
        if not self.current_student_id: return messagebox.showwarning("알림", "학생을 선택하세요.")
        weakness = get_latest_weakness(self.cfg, self.current_student_id)
        if not weakness: return messagebox.showwarning("경고", "먼저 분석을 수행하세요.")

        keywords = [kw.strip() for kw in weakness.split(",") if kw.strip()]
        found_pids = set()
        for kw in keywords:
            for p in search_passages(self.cfg, search=kw):
                found_pids.add(p["id"])

        if not found_pids:
            return messagebox.showinfo("결과", f"'{weakness}' 관련 문제가 DB에 없습니다.")

        # 문제 관리 탭의 장바구니에 추가
        bank_tab = self.app.bank_tab
        bank_tab._cart_clear()
        for pid in found_pids:
            bank_tab.cart[pid] = set()
        bank_tab._sync_cart_ui()

        messagebox.showinfo("추출 완료", f"{len(found_pids)}개 지문 추출.\n문제 관리 탭에서 생성하세요.")
        self.app.tabs.select(self.app.tab_bank)
