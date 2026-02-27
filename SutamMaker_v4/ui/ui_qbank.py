"""
SutamMaker v4 — ui/ui_qbank.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Question Bank DB browser window (multi-filter search, answer/meta editing).
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

from config import CIRCLE_NUMS, ALPHA_CHOICES
import exam_db


class QuestionBankWindow(tk.Toplevel):
    """Question-bank DB browser with multi-filter search, answer/explanation
    editing, metadata (difficulty / error-rate) editing, and exam generation."""

    def __init__(self, master, db_root, on_generate_callback=None):
        super().__init__(master)
        self.title("문제은행 데이터베이스 (다중 필터 + 메타데이터)")
        self.geometry("1300x850")
        self.db_root = db_root
        self.on_generate = on_generate_callback
        self.current_item = None
        self.answer_vars = []
        self.explanation_widgets = []
        self._editor_widgets = []
        self._meta_vars = []
        self._build_ui()
        self._refresh_list()

    # ── UI ──
    def _build_ui(self):
        # Top search bar (multi-filter)
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=6)

        tk.Label(top, text="검색 대상:", font=("맑은 고딕", 9)).pack(side="left")
        self.cb_field = ttk.Combobox(top, values=["전체", "태그", "본문"], width=8, state="readonly")
        self.cb_field.set("전체")
        self.cb_field.pack(side="left", padx=2)

        tk.Label(top, text="정답 유무:", font=("맑은 고딕", 9)).pack(side="left", padx=(10, 0))
        self.cb_ans_filter = ttk.Combobox(top, values=["전체", "정답 존재(O)", "정답 없음(X)"],
                                          width=12, state="readonly")
        self.cb_ans_filter.set("전체")
        self.cb_ans_filter.pack(side="left", padx=2)

        tk.Label(top, text="검색어:", font=("맑은 고딕", 9)).pack(side="left", padx=(10, 0))
        self.ent_search = tk.Entry(top, width=25, font=("맑은 고딕", 10))
        self.ent_search.pack(side="left", padx=4)
        self.ent_search.bind("<Return>", lambda e: self._do_search())

        tk.Button(top, text="적용", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._do_search).pack(side="left", padx=2)
        tk.Button(top, text="전체 보기", command=lambda: self._refresh_list()).pack(side="left", padx=2)
        tk.Button(top, text="인덱스 재생성", bg="#6366f1", fg="white",
                  font=("맑은 고딕", 9), command=self._rebuild_index).pack(side="left", padx=6)

        self.lbl_count = tk.Label(top, text="", font=("맑은 고딕", 9), fg="#64748b")
        self.lbl_count.pack(side="right", padx=8)

        # PanedWindow — left (list) / right (editor)
        pw = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashwidth=6)
        pw.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        # Left: Treeview
        left = tk.Frame(pw)
        pw.add(left, width=560, minsize=300)

        cols = ("ID", "제목", "태그", "정답", "난이도", "수정일", "문항수")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", height=28, selectmode="extended")
        for c, w in zip(cols, [90, 120, 90, 45, 45, 85, 45]):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, minwidth=35)
        vsb = tk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # Right: editor area
        right = tk.Frame(pw)
        pw.add(right, minsize=400)

        # Passage
        lf_pass = tk.LabelFrame(right, text=" 제시문 (Passage) ", font=("맑은 고딕", 9, "bold"))
        lf_pass.pack(fill="x", padx=6, pady=(4, 2))
        self.txt_passage = tk.Text(lf_pass, height=4, font=("Consolas", 9), bg="#f8fafc",
                                   state="disabled", wrap=tk.WORD)
        self.txt_passage.pack(fill="x", padx=4, pady=4)

        # Question editor (scrollable)
        lf_q = tk.LabelFrame(right, text=" 문항 편집 ", font=("맑은 고딕", 9, "bold"))
        lf_q.pack(fill="both", expand=True, padx=6, pady=2)

        canvas = tk.Canvas(lf_q, bg="#ffffff", highlightthickness=0)
        vsb2 = tk.Scrollbar(lf_q, orient="vertical", command=canvas.yview)
        self.q_frame = tk.Frame(canvas, bg="#ffffff")
        self.q_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.q_frame, anchor="nw")
        canvas.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # Tag display
        tag_row = tk.Frame(right)
        tag_row.pack(fill="x", padx=6, pady=2)
        tk.Label(tag_row, text="태그:", font=("맑은 고딕", 9, "bold")).pack(side="left")
        self.lbl_tags = tk.Label(tag_row, text="", font=("맑은 고딕", 9), fg="#059669")
        self.lbl_tags.pack(side="left", padx=4)

        # Bottom buttons
        bottom = tk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=8)

        tk.Button(bottom, text="저장", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=10, command=self._save_current).pack(side="left", padx=4)
        tk.Button(bottom, text="태그 수정", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=10, command=self._edit_tags).pack(side="left", padx=4)
        tk.Button(bottom, text="삭제", bg="#dc2626", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=10, command=self._delete_selected).pack(side="left", padx=4)
        tk.Button(bottom, text="시험지 생성", bg="#ea580c", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=14, command=self._generate_from_selected).pack(side="right", padx=4)
        tk.Button(bottom, text="JSON Export", bg="#0891b2", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=14, command=self._export_selected).pack(side="right", padx=4)

    # ── List (multi-filter) ──
    def _refresh_list(self, query=""):
        for c in self.tree.get_children():
            self.tree.delete(c)

        index = exam_db.load_index(self.db_root)
        field = self.cb_field.get() if hasattr(self, "cb_field") else "전체"
        ans_filter = self.cb_ans_filter.get() if hasattr(self, "cb_ans_filter") else "전체"
        query_lower = query.lower() if query else ""

        count = 0
        for iid, entry in index.get("items", {}).items():
            has_answer = entry.get("has_answer", False)
            if ans_filter == "정답 존재(O)" and not has_answer:
                continue
            if ans_filter == "정답 없음(X)" and has_answer:
                continue

            if query_lower:
                if field == "태그":
                    search_str = " ".join(entry.get("tags", [])).lower()
                elif field == "본문":
                    search_str = entry.get("passage", "").lower()
                else:
                    search_str = entry.get("search_text", "")
                if query_lower not in search_str:
                    continue

            self.tree.insert("", "end", iid=iid, values=(
                iid,
                entry.get("source_title", "")[:25],
                ", ".join(entry.get("tags", []))[:20],
                "O" if has_answer else "X",
                entry.get("difficulty", "3"),
                entry.get("updated_at", "")[:10],
                entry.get("question_count", 0),
            ))
            count += 1

        self.lbl_count.config(text=f"결과: {count}건")

    def _do_search(self):
        self._refresh_list(self.ent_search.get().strip())

    # ── Selection → editor ──
    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        item = exam_db.load_item(self.db_root, sel[0])
        if not item:
            return
        self.current_item = item
        self._load_editor(item)

    def _load_editor(self, item):
        # Passage
        self.txt_passage.config(state="normal")
        self.txt_passage.delete("1.0", tk.END)
        self.txt_passage.insert("1.0", item.get("passage", "") or "(없음)")
        self.txt_passage.config(state="disabled")

        # Tags
        self.lbl_tags.config(text=", ".join(item.get("tags", [])) or "(없음)")

        # Clear previous widgets
        for w in self._editor_widgets:
            w.destroy()
        self._editor_widgets.clear()
        self.answer_vars.clear()
        self.explanation_widgets.clear()
        self._meta_vars.clear()

        questions = item.get("questions", [])
        for idx, q in enumerate(questions):
            c_type = q.get("c_type", "num")
            frame = tk.Frame(self.q_frame, bg="#ffffff", bd=1, relief="groove")
            frame.pack(fill="x", padx=4, pady=4)
            self._editor_widgets.append(frame)

            # Header
            hdr = tk.Label(frame, text=f"문항 {idx+1}  (원본#{q.get('num', '?')})",
                           font=("맑은 고딕", 9, "bold"), bg="#eff6ff", anchor="w")
            hdr.pack(fill="x", padx=4, pady=(4, 2))

            # Body text
            txt_lbl = tk.Label(frame, text=q.get("text", "")[:200],
                               font=("맑은 고딕", 9), wraplength=550, justify="left",
                               bg="#ffffff", anchor="w")
            txt_lbl.pack(fill="x", padx=8, pady=2)

            # Choices
            choices = q.get("choices", [])
            for ci, ch in enumerate(choices):
                if ch:
                    if c_type == "alpha" and ci < len(ALPHA_CHOICES):
                        mark = ALPHA_CHOICES[ci]
                    elif ci < len(CIRCLE_NUMS):
                        mark = CIRCLE_NUMS[ci]
                    else:
                        mark = ""
                    clbl = tk.Label(frame, text=f"  {mark} {ch}",
                                    font=("맑은 고딕", 8), bg="#ffffff", anchor="w", fg="#475569")
                    clbl.pack(fill="x", padx=12)

            # Answer + metadata row
            ans_frame = tk.Frame(frame, bg="#ffffff")
            ans_frame.pack(fill="x", padx=8, pady=2)

            tk.Label(ans_frame, text="정답:", font=("맑은 고딕", 9, "bold"), bg="#ffffff").pack(side="left")
            if c_type == "alpha":
                ans_values = ["", "A", "B", "C", "D", "E"]
            else:
                ans_values = ["", "①", "②", "③", "④", "⑤"]
            ans_var = tk.StringVar(value=q.get("answer") or "")
            ttk.Combobox(ans_frame, textvariable=ans_var, values=ans_values,
                         width=6, state="readonly").pack(side="left", padx=4)
            self.answer_vars.append(ans_var)

            # Difficulty
            tk.Label(ans_frame, text="난이도(1~5):", font=("맑은 고딕", 8, "bold"),
                     bg="#ffffff").pack(side="left", padx=(12, 0))
            diff_var = tk.StringVar(value=str(q.get("difficulty", "3")))
            tk.Spinbox(ans_frame, from_=1, to=5, textvariable=diff_var, width=3,
                       font=("맑은 고딕", 9)).pack(side="left", padx=2)

            # Error rate
            tk.Label(ans_frame, text="오답률(%):", font=("맑은 고딕", 8, "bold"),
                     bg="#ffffff").pack(side="left", padx=(8, 0))
            err_var = tk.StringVar(value=str(q.get("error_rate", "0")))
            tk.Entry(ans_frame, textvariable=err_var, width=5,
                     font=("맑은 고딕", 9)).pack(side="left", padx=2)
            self._meta_vars.append({"diff": diff_var, "err": err_var})

            # Explanation
            tk.Label(ans_frame, text="해설:", font=("맑은 고딕", 9, "bold"),
                     bg="#ffffff").pack(side="left", padx=(12, 0))
            expl_txt = tk.Text(frame, height=2, font=("맑은 고딕", 9), wrap=tk.WORD)
            expl_txt.pack(fill="x", padx=8, pady=(0, 4))
            expl_txt.insert("1.0", q.get("explanation") or "")
            self.explanation_widgets.append(expl_txt)

    # ── Save (answer + explanation + metadata) ──
    def _save_current(self):
        if not self.current_item:
            return messagebox.showwarning("알림", "편집할 아이템을 선택하세요.", parent=self)

        questions = self.current_item.get("questions", [])
        for idx, q in enumerate(questions):
            if idx < len(self.answer_vars):
                raw_ans = self.answer_vars[idx].get()
                q["answer"] = exam_db.validate_answer(raw_ans, q.get("c_type", "num"))
                if raw_ans and q["answer"] is None and raw_ans.strip():
                    messagebox.showwarning("정답 오류",
                                           f"문항 {idx+1}: '{raw_ans}'은(는) 유효하지 않은 정답입니다.",
                                           parent=self)
            if idx < len(self.explanation_widgets):
                q["explanation"] = self.explanation_widgets[idx].get("1.0", tk.END).strip() or None
            if idx < len(self._meta_vars):
                q["difficulty"] = self._meta_vars[idx]["diff"].get()
                q["error_rate"] = self._meta_vars[idx]["err"].get()

        exam_db.save_item(self.db_root, self.current_item)
        exam_db.update_index_entry(self.db_root, self.current_item)
        self._refresh_list(self.ent_search.get().strip())
        messagebox.showinfo("저장", "정답/해설 및 메타데이터가 저장되었습니다.", parent=self)

    # ── Edit tags ──
    def _edit_tags(self):
        if not self.current_item:
            return messagebox.showwarning("알림", "아이템을 선택하세요.", parent=self)
        current = ", ".join(self.current_item.get("tags", []))
        new_tags = simpledialog.askstring("태그 수정", f"태그 (쉼표 구분):\n현재: {current}",
                                          initialvalue=current, parent=self)
        if new_tags is None:
            return
        self.current_item["tags"] = [t.strip() for t in new_tags.split(",") if t.strip()]
        exam_db.save_item(self.db_root, self.current_item)
        exam_db.update_index_entry(self.db_root, self.current_item)
        self.lbl_tags.config(text=", ".join(self.current_item["tags"]) or "(없음)")
        self._refresh_list(self.ent_search.get().strip())

    # ── Delete ──
    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("알림", "삭제할 아이템을 선택하세요.", parent=self)
        if not messagebox.askyesno("확인", f"{len(sel)}개 아이템을 삭제하시겠습니까?", parent=self):
            return
        for iid in sel:
            exam_db.delete_item(self.db_root, iid)
            exam_db.remove_index_entry(self.db_root, iid)
        self.current_item = None
        self._refresh_list(self.ent_search.get().strip())

    # ── Generate exam from selected ──
    def _generate_from_selected(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("알림", "시험지를 생성할 아이템을 선택하세요.", parent=self)
        items = [exam_db.load_item(self.db_root, iid) for iid in sel]
        items = [it for it in items if it]
        if not items:
            return
        if self.on_generate:
            self.on_generate(items)
            self.destroy()

    # ── JSON Export ──
    def _export_selected(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("알림", "Export할 아이템을 선택하세요.", parent=self)
        items = [exam_db.load_item(self.db_root, iid) for iid in sel]
        items = [it for it in items if it]
        if not items:
            return
        target_dir = filedialog.askdirectory(title="Export 저장 폴더", parent=self)
        if not target_dir:
            return
        title = simpledialog.askstring("제목", "Export 파일명:", initialvalue="exam_export", parent=self)
        if not title:
            return
        path = exam_db.export_from_db_items(items, target_dir, title, display_title=title)
        messagebox.showinfo("Export 완료", f"저장: {path}", parent=self)

    # ── Rebuild index ──
    def _rebuild_index(self):
        idx = exam_db.build_index(self.db_root)
        cnt = len(idx.get("items", {}))
        self._refresh_list(self.ent_search.get().strip())
        messagebox.showinfo("완료", f"인덱스 재생성: {cnt}개 아이템", parent=self)
