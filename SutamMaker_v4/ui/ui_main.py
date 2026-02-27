"""
SutamMaker v4 — ui/ui_main.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Main application window (SutamMakerApp).

Changes from original draft
----------------------------
- Import paths unified to package-relative (``from config import …``)
- Function names aligned to data_manager public API (no underscore prefix)
- ``convert_docx_pairs_to_pdf`` imported from ``core.docx_generator``
- Inline imports (docx, subprocess, json) moved to module top
- ``_delete_selected_bank`` uses ``save_qbank_data()`` instead of raw file write
- Redundant list copy in ``_generate`` removed
- ``_update_status`` debounced (500 ms); bank/history counts cached
- ``_on_file_drop`` uses top-level docx import
"""

import os
import re
import random
import datetime
import platform
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

import docx as _docx_mod  # only used for Drag-and-Drop .docx reading

try:
    from tkinterdnd2 import DND_FILES
except ImportError:
    DND_FILES = None

from config import (
    SCRIPT_DIR, CONFIG_FILE, QBANK_FILE,
    SUBJECT_LIST, DIFFICULTY_LIST, AI_MASTER_PROMPT_TEMPLATE,
    CIRCLE_NUMS,
    load_config, save_config,
)
from core.parser_engine import parse_exam_text, parse_answer_by_question
from core.docx_generator import create_exam_docx, create_answer_docx, convert_docx_pairs_to_pdf
from data.data_manager import (
    save_to_qbank, load_qbank, clear_qbank, save_qbank_data,
    add_history, load_history,
    save_draft, load_draft, clear_draft,
)
from ui.ui_qbank import QuestionBankWindow
import exam_db


class SutamMakerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("사탐/경제 모의고사 자동 생성기 PRO v4.0")
        self.root.geometry("1100x980")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.cfg = load_config()
        self._status_timer = None       # debounce timer for _update_status
        self._bank_count = 0            # cached qbank count
        self._hist_count = 0            # cached history count

        self._build_ui()
        self._bind_shortcuts()
        self._restore_draft()
        self._start_autosave()

        # Drag & Drop (requires tkinterdnd2)
        if DND_FILES and hasattr(self.root, "drop_target_register"):
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_file_drop)
            except Exception:
                pass

    # ─────────────────────── UI ───────────────────────
    def _build_ui(self):
        # Status bar
        self.status_bar = tk.Frame(self.root, bg="#e2e8f0", height=26)
        self.status_bar.pack(fill="x", side="bottom")
        self.status_bar.pack_propagate(False)
        self.lbl_status = tk.Label(self.status_bar, text="", font=("맑은 고딕", 8),
                                   bg="#e2e8f0", fg="#475569")
        self.lbl_status.pack(side="left", padx=10)

        # Tabs
        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=(8, 4))

        self.tab_main = ttk.Frame(self.tabs)
        self.tab_preview = ttk.Frame(self.tabs)
        self.tab_qbank = ttk.Frame(self.tabs)
        self.tab_history = ttk.Frame(self.tabs)
        self.tab_ai = ttk.Frame(self.tabs)
        self.tab_db = ttk.Frame(self.tabs)

        self.tabs.add(self.tab_main, text="  시험지 생성  ")
        self.tabs.add(self.tab_preview, text="  미리보기  ")
        self.tabs.add(self.tab_qbank, text="  문제 은행  ")
        self.tabs.add(self.tab_history, text="  생성 히스토리  ")
        self.tabs.add(self.tab_ai, text="  AI 명령어  ")
        self.tabs.add(self.tab_db, text="  DB 도구  ")

        self._build_main_tab()
        self._build_preview_tab()
        self._build_qbank_tab()
        self._build_history_tab()
        self._build_ai_tab()
        self._build_db_tab()
        self._update_status()

    def _bind_shortcuts(self):
        self.root.bind("<Control-g>", lambda e: self._generate("both"))
        self.root.bind("<Control-p>", lambda e: self._show_preview())
        self.root.bind("<Control-s>", lambda e: self._save_draft_now())

    # ─────────────── TAB 1: 시험지 생성 ───────────────
    def _build_main_tab(self):
        f_cfg = tk.LabelFrame(self.tab_main, text=" 설정 ", font=("맑은 고딕", 10, "bold"))
        f_cfg.pack(fill="x", padx=12, pady=6)

        r0 = tk.Frame(f_cfg); r0.pack(fill="x", padx=8, pady=3)
        tk.Label(r0, text="저장 폴더:", width=12, anchor="w").pack(side="left")
        self.dir_var = tk.StringVar(value=self.cfg.get("last_dir", SCRIPT_DIR))
        tk.Entry(r0, textvariable=self.dir_var, width=42).pack(side="left", padx=4)
        tk.Button(r0, text="찾아보기", command=self._browse_dir).pack(side="left", padx=2)
        tk.Button(r0, text="폴더 열기", command=self._open_target_dir).pack(side="left", padx=2)

        r1 = tk.Frame(f_cfg); r1.pack(fill="x", padx=8, pady=3)
        tk.Label(r1, text="배너 로고:", width=12, anchor="w").pack(side="left")
        self.logo_var = tk.StringVar(value=self.cfg.get("last_logo", ""))
        tk.Entry(r1, textvariable=self.logo_var, width=42).pack(side="left", padx=4)
        tk.Button(r1, text="이미지 선택", command=self._browse_logo).pack(side="left", padx=2)
        tk.Button(r1, text="초기화", command=lambda: self.logo_var.set("")).pack(side="left", padx=2)

        r2 = tk.Frame(f_cfg); r2.pack(fill="x", padx=8, pady=3)
        tk.Label(r2, text="파일 제목:", width=12, anchor="w").pack(side="left")
        self.ent_title = tk.Entry(r2, width=30)
        self.ent_title.pack(side="left", padx=4)
        tk.Label(r2, text="(예: 이준호_경제)", fg="gray", font=("맑은 고딕", 8)).pack(side="left")

        r3 = tk.Frame(f_cfg); r3.pack(fill="x", padx=8, pady=3)
        tk.Label(r3, text="폰트:", width=12, anchor="w").pack(side="left")
        self.font_combo = ttk.Combobox(r3, values=["맑은 고딕", "바탕", "굴림", "나눔고딕"],
                                       width=14, state="readonly")
        self.font_combo.set(self.cfg.get("font_name", "맑은 고딕"))
        self.font_combo.pack(side="left", padx=4)
        tk.Label(r3, text="크기:").pack(side="left")
        self.size_combo = ttk.Combobox(r3, values=["9", "10", "11", "12"], width=4, state="readonly")
        self.size_combo.set(self.cfg.get("font_size", "10"))
        self.size_combo.pack(side="left", padx=4)

        r4 = tk.Frame(f_cfg); r4.pack(fill="x", padx=8, pady=3)
        tk.Label(r4, text="배치 생성:", width=12, anchor="w").pack(side="left")
        self.ent_batch = tk.Entry(r4, width=30)
        self.ent_batch.pack(side="left", padx=4)
        tk.Label(r4, text="(쉼표로 구분: 이준호,김철수,박영희)", fg="gray",
                 font=("맑은 고딕", 8)).pack(side="left")

        r5 = tk.Frame(f_cfg); r5.pack(fill="x", padx=8, pady=3)
        self.var_shuffle = tk.BooleanVar(value=False)
        tk.Checkbutton(r5, text="문제 순서 섞기 (학생마다 랜덤)",
                       variable=self.var_shuffle).pack(side="left", padx=(88, 10))
        tk.Label(r5, text="워터마크:").pack(side="left")
        self.ent_watermark = tk.Entry(r5, width=15)
        self.ent_watermark.pack(side="left", padx=4)

        # Exam text
        tk.Label(self.tab_main, text="[1] 문제 입력 (AI 출력물 붙여넣기 가능)",
                 font=("맑은 고딕", 10, "bold"), fg="#2563eb").pack(anchor="w", padx=12, pady=(8, 0))
        self.txt_exam = tk.Text(self.tab_main, wrap=tk.WORD, height=14, font=("Consolas", 10))
        self.txt_exam.pack(fill="both", expand=True, padx=12, pady=4)
        self.txt_exam.bind("<KeyRelease>", self._on_text_change)

        # Answer text
        tk.Label(self.tab_main, text="[2] 정답 및 해설 (선택 - AI가 생성한 정답/해설 붙여넣기)",
                 font=("맑은 고딕", 10, "bold"), fg="#16a34a").pack(anchor="w", padx=12, pady=(4, 0))
        self.txt_ans = tk.Text(self.tab_main, wrap=tk.WORD, height=7, font=("Consolas", 10))
        self.txt_ans.pack(fill="x", padx=12, pady=4)

        # Buttons
        bf = tk.Frame(self.tab_main); bf.pack(pady=8)
        tk.Button(bf, text="Word만", bg="#0891b2", fg="white", font=("맑은 고딕", 10, "bold"),
                  width=14, height=2, command=lambda: self._generate("word")).pack(side="left", padx=6)
        tk.Button(bf, text="PDF만", bg="#dc2626", fg="white", font=("맑은 고딕", 10, "bold"),
                  width=14, height=2, command=lambda: self._generate("pdf")).pack(side="left", padx=6)
        tk.Button(bf, text="Word + PDF", bg="#ea580c", fg="white", font=("맑은 고딕", 10, "bold"),
                  width=18, height=2, command=lambda: self._generate("both")).pack(side="left", padx=6)
        tk.Button(bf, text="미리보기 (Ctrl+P)", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"),
                  width=18, height=2, command=self._show_preview).pack(side="left", padx=6)

    # ─────────────── TAB 2: 미리보기 ───────────────
    def _build_preview_tab(self):
        ctrl = tk.Frame(self.tab_preview); ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="파싱 새로고침", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._show_preview).pack(side="left", padx=4)
        tk.Button(ctrl, text="문제 은행에 저장 (정답 포함)", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._save_to_bank_from_preview).pack(side="left", padx=4)
        self.lbl_preview_count = tk.Label(ctrl, text="", font=("맑은 고딕", 9), fg="#64748b")
        self.lbl_preview_count.pack(side="right", padx=8)

        self.txt_preview = tk.Text(self.tab_preview, font=("Consolas", 10), bg="#f8fafc", state="disabled")
        scroll = tk.Scrollbar(self.tab_preview, command=self.txt_preview.yview)
        self.txt_preview.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.txt_preview.pack(fill="both", expand=True, padx=12, pady=6)

    def _show_preview(self, event=None):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showwarning("경고", "문제 텍스트를 먼저 입력하세요.")
            return

        parsed = parse_exam_text(raw)
        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", tk.END)

        if not parsed:
            first_lines = raw[:300].replace('\r', '\n')
            self.txt_preview.insert(tk.END,
                "파싱된 문제가 없습니다.\n\n"
                "── 지원 형식 ──\n"
                "  01 문제 본문...\n"
                "  01. 문제 본문...\n"
                "  1) 문제 본문...\n\n"
                "── 입력된 텍스트 (처음 300자) ──\n"
                f"{first_lines}\n"
            )
            self.txt_preview.config(state="disabled")
            return

        for i, q in enumerate(parsed, 1):
            self.txt_preview.insert(tk.END, f"{'='*50}\n")
            self.txt_preview.insert(tk.END, f"  문항 {i} (원본 번호: {q['num']})\n")
            self.txt_preview.insert(tk.END, f"{'='*50}\n\n")
            self.txt_preview.insert(tk.END, f"[본문]\n{q['text']}\n\n")
            if q.get("jesi"):
                self.txt_preview.insert(tk.END, f"[제시문]\n{q['jesi']}\n\n")
            if q.get("bogi"):
                self.txt_preview.insert(tk.END, f"[보기]\n{q['bogi']}\n\n")
            if q.get("tables"):
                self.txt_preview.insert(tk.END, f"[표] {len(q['tables'])}개 감지됨\n\n")
            if q.get("image_path"):
                self.txt_preview.insert(tk.END, f"[이미지] {q['image_path']}\n\n")
            if q.get("graph_tag"):
                self.txt_preview.insert(tk.END, f"[그래프 태그]\n{q['graph_tag']}\n\n")
            if q.get("choices"):
                self.txt_preview.insert(tk.END, "[선택지]\n")
                for j, ch in enumerate(q["choices"]):
                    if ch:
                        self.txt_preview.insert(tk.END, f"  {CIRCLE_NUMS[j]} {ch}\n")
                self.txt_preview.insert(tk.END, "\n")

        self.txt_preview.config(state="disabled")
        self.lbl_preview_count.config(text=f"총 {len(parsed)}문항 파싱됨")
        self.tabs.select(self.tab_preview)
        self._set_status(f"미리보기: {len(parsed)}문항 파싱 완료")

    def _save_to_bank_from_preview(self):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        if not raw:
            return messagebox.showwarning("경고", "문제 텍스트가 없습니다.")
        parsed = parse_exam_text(raw)
        if not parsed:
            return messagebox.showwarning("오류", "파싱된 문제가 없습니다.")

        raw_ans = self.txt_ans.get("1.0", tk.END).strip()
        answers_map = parse_answer_by_question(raw_ans)
        tag = simpledialog.askstring("태그", "저장 태그 (예: 경제_중간고사):", parent=self.root)
        if tag is None:
            return

        cnt = save_to_qbank(parsed, tag=tag or "", answers_map=answers_map)
        messagebox.showinfo("완료",
            f"{cnt}문항을 문제 은행에 저장했습니다.\n"
            f"(정답/해설 {'포함' if answers_map else '없음'})")
        self._refresh_qbank()

    # ─────────────── TAB 3: 문제 은행 ───────────────
    def _build_qbank_tab(self):
        ctrl = tk.Frame(self.tab_qbank); ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="새로고침", command=self._refresh_qbank).pack(side="left", padx=4)
        tk.Button(ctrl, text="선택 문항 -> 시험지+답안지", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._load_from_bank).pack(side="left", padx=4)
        tk.Button(ctrl, text="전체 선택", bg="#6366f1", fg="white",
                  font=("맑은 고딕", 9), command=self._select_all_bank).pack(side="left", padx=4)
        tk.Button(ctrl, text="선택 해제", bg="#94a3b8", fg="white",
                  font=("맑은 고딕", 9), command=self._deselect_all_bank).pack(side="left", padx=4)
        tk.Button(ctrl, text="전체 삭제", bg="#dc2626", fg="white",
                  command=self._clear_bank).pack(side="right", padx=4)
        tk.Button(ctrl, text="선택 항목 삭제", bg="#f97316", fg="white",
                  command=self._delete_selected_bank).pack(side="right", padx=4)
        self.lbl_bank_count = tk.Label(ctrl, text="", font=("맑은 고딕", 9), fg="#64748b")
        self.lbl_bank_count.pack(side="right", padx=8)

        tk.Label(self.tab_qbank,
                 text="서로 다른 세트에서 자유롭게 골라 조합하세요. 번호는 자동으로 1번부터 재매김됩니다. 정답/해설도 자동 매칭!",
                 font=("맑은 고딕", 9), fg="#059669", bg="#ecfdf5", relief="groove",
                 padx=8, pady=4).pack(fill="x", padx=12, pady=(0, 4))

        cols = ("번호", "태그", "정답유무", "저장일", "내용 미리보기")
        self.tree_bank = ttk.Treeview(self.tab_qbank, columns=cols, show="headings",
                                      height=18, selectmode="extended")
        for c, w in zip(cols, [50, 100, 70, 110, 450]):
            self.tree_bank.heading(c, text=c)
            self.tree_bank.column(c, width=w)
        self.tree_bank.pack(fill="both", expand=True, padx=12, pady=6)

        self.lbl_selected = tk.Label(self.tab_qbank, text="선택된 문항: 0개",
                                     font=("맑은 고딕", 10, "bold"), fg="#2563eb")
        self.lbl_selected.pack(anchor="w", padx=12, pady=(4, 2))
        self.tree_bank.bind("<<TreeviewSelect>>", self._on_bank_select)
        self._refresh_qbank()

    def _on_bank_select(self, event=None):
        self.lbl_selected.config(text=f"선택된 문항: {len(self.tree_bank.selection())}개")

    def _select_all_bank(self):
        self.tree_bank.selection_set(self.tree_bank.get_children())
        self._on_bank_select()

    def _deselect_all_bank(self):
        self.tree_bank.selection_remove(self.tree_bank.selection())
        self._on_bank_select()

    def _refresh_qbank(self):
        for i in self.tree_bank.get_children():
            self.tree_bank.delete(i)
        bank = load_qbank()
        for i, q in enumerate(bank, 1):
            preview = q.get("text", "")[:60].replace("\n", " ")
            saved = q.get("saved_at", "")[:10]
            tag = q.get("tag", "")
            has_ans = "O" if q.get("answer", "").strip() else "X"
            self.tree_bank.insert("", "end", iid=str(i - 1),
                                  values=(i, tag, has_ans, saved, preview))
        self._bank_count = len(bank)
        self.lbl_bank_count.config(text=f"총 {self._bank_count}문항")

    def _load_from_bank(self):
        sel = self.tree_bank.selection()
        if not sel:
            return messagebox.showwarning("알림", "문항을 선택하세요. (Ctrl+클릭으로 여러 개 선택 가능)")

        bank = load_qbank()
        exam_lines, ans_lines = [], []
        new_num = 1
        for iid in sel:
            idx = int(iid)
            if idx >= len(bank):
                continue
            q = bank[idx]

            # Rebuild exam text
            body = f"{new_num}. {q['text']}"
            if q.get("jesi"):
                body += f"\n<제시문>{q['jesi']}</제시문>"
            if q.get("graph_tag"):
                body += f"\n{q['graph_tag']}"
            if q.get("bogi"):
                body += f"\n<보기>\n{q['bogi']}"
            if q.get("choices"):
                for j, ch in enumerate(q["choices"]):
                    if ch:
                        body += f"\n{CIRCLE_NUMS[j]} {ch}"
            exam_lines.append(body)

            # Rebuild answer text with renumbered question
            ans_raw = q.get("answer", "").strip()
            if ans_raw:
                old_num = q.get("num", "")
                renamed = re.sub(
                    rf"^\s*{re.escape(old_num)}\s*(?:번|[\.\)])",
                    f"{new_num}번", ans_raw, count=1, flags=re.MULTILINE)
                renamed = re.sub(
                    rf"\b{re.escape(old_num)}\s*(?:번\s*)?정답",
                    f"{new_num}번 정답", renamed, flags=re.IGNORECASE)
                ans_lines.append(renamed)
            else:
                ans_lines.append(f"{new_num}번 정답: (정보 없음)")
            new_num += 1

        exam_text = "\n\n".join(exam_lines)
        ans_text = "\n\n".join(ans_lines)

        current_exam = self.txt_exam.get("1.0", tk.END).strip()
        if current_exam:
            choice = messagebox.askyesnocancel(
                "불러오기 방법",
                f"선택한 {len(sel)}문항을 불러옵니다.\n\n"
                "[예] 기존 내용 뒤에 추가\n"
                "[아니오] 기존 내용 지우고 새로 입력\n"
                "[취소] 취소")
            if choice is None:
                return
            if choice:
                self.txt_exam.insert(tk.END, "\n\n" + exam_text)
                self.txt_ans.insert(tk.END, "\n\n" + ans_text)
            else:
                self.txt_exam.delete("1.0", tk.END)
                self.txt_exam.insert("1.0", exam_text)
                self.txt_ans.delete("1.0", tk.END)
                self.txt_ans.insert("1.0", ans_text)
        else:
            self.txt_exam.insert("1.0", exam_text)
            self.txt_ans.insert("1.0", ans_text)

        self.tabs.select(self.tab_main)
        self._set_status(f"문제 은행에서 {len(sel)}문항 + 정답/해설 불러옴")

    def _delete_selected_bank(self):
        sel = self.tree_bank.selection()
        if not sel:
            return messagebox.showwarning("알림", "삭제할 문항을 선택하세요.")
        if not messagebox.askyesno("확인", f"선택한 {len(sel)}개 문항을 삭제하시겠습니까?"):
            return
        bank = load_qbank()
        for idx in sorted((int(iid) for iid in sel), reverse=True):
            if idx < len(bank):
                bank.pop(idx)
        save_qbank_data(bank)
        self._refresh_qbank()

    def _clear_bank(self):
        if not messagebox.askyesno("확인", "문제 은행을 전체 삭제하시겠습니까?"):
            return
        clear_qbank()
        self._refresh_qbank()

    # ─────────────── TAB 4: 생성 히스토리 ───────────────
    def _build_history_tab(self):
        ctrl = tk.Frame(self.tab_history); ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="새로고침", command=self._refresh_history).pack(side="left", padx=4)
        tk.Button(ctrl, text="폴더 열기", command=self._open_history_dir).pack(side="left", padx=4)

        cols = ("시각", "제목", "학생", "문항수", "형식", "경로")
        self.tree_hist = ttk.Treeview(self.tab_history, columns=cols, show="headings", height=18)
        for c, w in zip(cols, [130, 140, 100, 60, 70, 350]):
            self.tree_hist.heading(c, text=c)
            self.tree_hist.column(c, width=w)
        self.tree_hist.pack(fill="both", expand=True, padx=12, pady=6)
        self._refresh_history()

    def _refresh_history(self):
        for i in self.tree_hist.get_children():
            self.tree_hist.delete(i)
        history = load_history()
        for h in history:
            self.tree_hist.insert("", "end", values=(
                h.get("timestamp", "")[:19].replace("T", " "),
                h.get("title", ""),
                h.get("student", ""),
                h.get("q_count", ""),
                h.get("format", ""),
                h.get("path", ""),
            ))
        self._hist_count = len(history)

    def _open_history_dir(self):
        sel = self.tree_hist.selection()
        if sel:
            vals = self.tree_hist.item(sel[0])["values"]
            path = vals[-1] if vals else ""
            if path and os.path.isdir(str(path)):
                self._open_dir(str(path))
                return
        self._open_dir(self.dir_var.get())

    # ─────────────── TAB 5: AI 명령어 ───────────────
    def _build_ai_tab(self):
        f_settings = tk.LabelFrame(self.tab_ai, text=" AI 프롬프트 설정 ",
                                   font=("맑은 고딕", 10, "bold"))
        f_settings.pack(fill="x", padx=12, pady=6)

        r0 = tk.Frame(f_settings); r0.pack(fill="x", padx=8, pady=3)
        tk.Label(r0, text="과목:", width=10, anchor="w").pack(side="left")
        self.ai_subject = ttk.Combobox(r0, values=SUBJECT_LIST, width=20, state="readonly")
        self.ai_subject.set("경제")
        self.ai_subject.pack(side="left", padx=4)

        r1 = tk.Frame(f_settings); r1.pack(fill="x", padx=8, pady=3)
        tk.Label(r1, text="난이도:", width=10, anchor="w").pack(side="left")
        self.ai_difficulty = ttk.Combobox(r1, values=DIFFICULTY_LIST, width=40, state="readonly")
        self.ai_difficulty.set(DIFFICULTY_LIST[3])
        self.ai_difficulty.pack(side="left", padx=4)

        r2 = tk.Frame(f_settings); r2.pack(fill="x", padx=8, pady=3)
        tk.Label(r2, text="문항 수:", width=10, anchor="w").pack(side="left")
        self.ai_num_q = ttk.Combobox(r2, values=["5", "10", "15", "20", "25", "30"], width=8)
        self.ai_num_q.set("10")
        self.ai_num_q.pack(side="left", padx=4)

        r3 = tk.Frame(f_settings); r3.pack(fill="x", padx=8, pady=3)
        tk.Label(r3, text="범위/단원:", width=10, anchor="w").pack(side="left")
        self.ai_scope = tk.Entry(r3, width=50)
        self.ai_scope.pack(side="left", padx=4)
        tk.Label(r3, text="(예: 수요공급, 시장균형, 비교우위)", fg="gray",
                 font=("맑은 고딕", 8)).pack(side="left")

        r4 = tk.Frame(f_settings); r4.pack(fill="x", padx=8, pady=3)
        tk.Label(r4, text="추가 지시:", width=10, anchor="w").pack(side="left")
        self.ai_extra = tk.Entry(r4, width=50)
        self.ai_extra.pack(side="left", padx=4)
        tk.Label(r4, text="(예: 그래프 문항 3개 포함)", fg="gray",
                 font=("맑은 고딕", 8)).pack(side="left")

        bf = tk.Frame(self.tab_ai); bf.pack(fill="x", padx=12, pady=6)
        tk.Button(bf, text="AI 명령어 생성", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=18, height=2,
                  command=self._generate_ai_prompt).pack(side="left", padx=6)
        tk.Button(bf, text="클립보드에 복사", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=18, height=2,
                  command=self._copy_ai_prompt).pack(side="left", padx=6)
        tk.Button(bf, text="AI 결과 -> 시험지 탭으로", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=22, height=2,
                  command=self._paste_ai_result).pack(side="left", padx=6)

        info_frame = tk.Frame(self.tab_ai, bg="#f0f9ff", relief="groove", bd=1)
        info_frame.pack(fill="x", padx=12, pady=(0, 4))
        tk.Label(info_frame,
                 text="사용법: [AI 명령어 생성] -> [클립보드에 복사] -> ChatGPT/Claude에 붙여넣기 -> AI 출력 복사 -> [AI 결과 -> 시험지 탭으로]",
                 font=("맑은 고딕", 9), fg="#1e40af", bg="#f0f9ff", wraplength=900,
                 justify="left").pack(padx=8, pady=4)

        tk.Label(self.tab_ai, text="생성된 AI 명령어 미리보기:",
                 font=("맑은 고딕", 10, "bold"), fg="#7c3aed").pack(anchor="w", padx=12, pady=(4, 0))
        self.txt_ai_prompt = tk.Text(self.tab_ai, wrap=tk.WORD, height=20,
                                     font=("Consolas", 9), bg="#faf5ff")
        scroll_ai = tk.Scrollbar(self.tab_ai, command=self.txt_ai_prompt.yview)
        self.txt_ai_prompt.configure(yscrollcommand=scroll_ai.set)
        scroll_ai.pack(side="right", fill="y", padx=(0, 12))
        self.txt_ai_prompt.pack(fill="both", expand=True, padx=12, pady=4)

    def _generate_ai_prompt(self):
        prompt = AI_MASTER_PROMPT_TEMPLATE.format(
            subject=self.ai_subject.get(),
            difficulty=self.ai_difficulty.get(),
            num_questions=self.ai_num_q.get(),
            scope=self.ai_scope.get().strip() or "전 범위",
            extra=self.ai_extra.get().strip() or "없음",
        )
        self.txt_ai_prompt.delete("1.0", tk.END)
        self.txt_ai_prompt.insert("1.0", prompt)
        self._set_status(f"AI 명령어 생성 완료 ({self.ai_subject.get()}, {self.ai_num_q.get()}문항)")

    def _copy_ai_prompt(self):
        prompt = self.txt_ai_prompt.get("1.0", tk.END).strip()
        if not prompt:
            self._generate_ai_prompt()
            prompt = self.txt_ai_prompt.get("1.0", tk.END).strip()
        if prompt:
            self.root.clipboard_clear()
            self.root.clipboard_append(prompt)
            self.root.update()
            messagebox.showinfo("복사 완료",
                                "AI 명령어가 클립보드에 복사되었습니다!\n\n"
                                "ChatGPT 또는 Claude에 붙여넣기(Ctrl+V)하세요.")

    def _paste_ai_result(self):
        try:
            clipboard = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("경고", "클립보드가 비어있습니다.\nAI 출력을 먼저 복사하세요.")
            return
        if not clipboard.strip():
            messagebox.showwarning("경고", "클립보드가 비어있습니다.")
            return

        split_patterns = [
            r"\n\s*\[출력\s*2\s*[—\-]\s*교사용",
            r"\n\s*#{1,3}\s*(?:정답\s*및\s*해설|교사용\s*정답)",
            r"\n\s*(?:정답\s*및\s*해설|교사용\s*정답\s*및\s*해설)\s*\n",
            r"\n\s*[-=]{5,}\s*\n\s*(?:정답|해설)",
        ]
        exam_part, ans_part = clipboard, ""
        for pat in split_patterns:
            m = re.search(pat, clipboard, re.IGNORECASE)
            if m:
                exam_part = clipboard[:m.start()].strip()
                ans_part = clipboard[m.start():].strip()
                break

        self.txt_exam.delete("1.0", tk.END)
        self.txt_exam.insert("1.0", exam_part)
        if ans_part:
            self.txt_ans.delete("1.0", tk.END)
            self.txt_ans.insert("1.0", ans_part)

        self.tabs.select(self.tab_main)
        parsed = parse_exam_text(exam_part)
        self._set_status(
            f"AI 결과 붙여넣기 완료: {len(parsed)}문항 감지, "
            f"답안지 {'있음' if ans_part else '없음'}")

    # ─────────────── TAB 6: DB 도구 ───────────────
    def _build_db_tab(self):
        lf_folder = tk.LabelFrame(self.tab_db, text=" DB 폴더 설정 ",
                                  font=("맑은 고딕", 10, "bold"))
        lf_folder.pack(fill="x", padx=12, pady=6)
        r0 = tk.Frame(lf_folder); r0.pack(fill="x", padx=8, pady=4)
        tk.Label(r0, text="DB 폴더:", width=10, anchor="w").pack(side="left")
        self.db_dir_var = tk.StringVar(
            value=self.cfg.get("db_root", os.path.join(SCRIPT_DIR, "db")))
        tk.Entry(r0, textvariable=self.db_dir_var, width=45).pack(side="left", padx=4)
        tk.Button(r0, text="찾아보기", command=self._browse_db_dir).pack(side="left", padx=2)

        lf_actions = tk.LabelFrame(self.tab_db, text=" DB 작업 ", font=("맑은 고딕", 10, "bold"))
        lf_actions.pack(fill="x", padx=12, pady=6)
        bf = tk.Frame(lf_actions); bf.pack(padx=8, pady=8)
        for txt, bg, cmd in [
            ("DB 저장\n(파싱 결과 저장)", "#16a34a", self._save_to_db),
            ("문제은행\n(브라우저 열기)", "#2563eb", self._open_qbank_window),
            ("DB 백업\n(zip 생성)", "#7c3aed", self._backup_db),
            ("DB 복원\n(zip에서)", "#dc2626", self._restore_db),
            ("JSON Export\n(시험지 JSON)", "#0891b2", self._export_json),
        ]:
            tk.Button(bf, text=txt, bg=bg, fg="white", font=("맑은 고딕", 10, "bold"),
                      width=16, height=3, command=cmd).pack(side="left", padx=6)

        # Header/footer settings
        lf_hf = tk.LabelFrame(self.tab_db, text=" 머리글 / 바닥글 설정 (DOCX 적용) ",
                               font=("맑은 고딕", 10, "bold"))
        lf_hf.pack(fill="x", padx=12, pady=6)

        r1 = tk.Frame(lf_hf); r1.pack(fill="x", padx=8, pady=2)
        tk.Label(r1, text="학원명:", width=10, anchor="w").pack(side="left")
        self.hf_academy = tk.Entry(r1, width=30)
        self.hf_academy.pack(side="left", padx=4)
        self.hf_academy.insert(0, self.cfg.get("academy_name", ""))
        tk.Label(r1, text="시험명:", width=8, anchor="w").pack(side="left", padx=(8, 0))
        self.hf_exam = tk.Entry(r1, width=25)
        self.hf_exam.pack(side="left", padx=4)
        self.hf_exam.insert(0, self.cfg.get("exam_name", ""))

        r2 = tk.Frame(lf_hf); r2.pack(fill="x", padx=8, pady=2)
        tk.Label(r2, text="연락처:", width=10, anchor="w").pack(side="left")
        self.hf_contact = tk.Entry(r2, width=30)
        self.hf_contact.pack(side="left", padx=4)
        self.hf_contact.insert(0, self.cfg.get("contact_info", ""))
        tk.Label(r2, text="저작권:", width=8, anchor="w").pack(side="left", padx=(8, 0))
        self.hf_copyright = tk.Entry(r2, width=25)
        self.hf_copyright.pack(side="left", padx=4)
        self.hf_copyright.insert(0, self.cfg.get("copyright_text", ""))

        r3 = tk.Frame(lf_hf); r3.pack(fill="x", padx=8, pady=4)
        self.hf_page_var = tk.BooleanVar(value=self.cfg.get("show_page_number", False))
        tk.Checkbutton(r3, text="페이지 번호 표시 (Page X)",
                       variable=self.hf_page_var).pack(side="left", padx=(76, 10))
        self.hf_date_var = tk.BooleanVar(value=self.cfg.get("show_date", False))
        tk.Checkbutton(r3, text="날짜 표시", variable=self.hf_date_var).pack(side="left", padx=10)
        tk.Button(r3, text="설정 저장", bg="#475569", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._save_hf_config).pack(side="right", padx=8)

        tk.Label(self.tab_db,
                 text="머리글 예시: 학원명 | 시험명 | 날짜 | Page X\n"
                      "바닥글 예시: 저작권 | 연락처 | Page X\n"
                      "(설정 저장 후 시험지 생성 시 자동 적용됩니다)",
                 font=("맑은 고딕", 9), fg="#64748b", bg="#f1f5f9", relief="groove",
                 padx=8, pady=6, justify="left").pack(fill="x", padx=12, pady=(0, 6))

    # ─── DB helpers ───
    def _get_db_root(self):
        return self.db_dir_var.get()

    def _browse_db_dir(self):
        d = filedialog.askdirectory(title="DB 폴더 선택")
        if d:
            self.db_dir_var.set(d)
            self.cfg["db_root"] = d
            save_config(self.cfg)

    def _save_hf_config(self):
        self.cfg["academy_name"] = self.hf_academy.get().strip()
        self.cfg["exam_name"] = self.hf_exam.get().strip()
        self.cfg["contact_info"] = self.hf_contact.get().strip()
        self.cfg["copyright_text"] = self.hf_copyright.get().strip()
        self.cfg["show_page_number"] = self.hf_page_var.get()
        self.cfg["show_date"] = self.hf_date_var.get()
        self.cfg["db_root"] = self.db_dir_var.get()
        save_config(self.cfg)
        self._set_status("머리글/바닥글 설정 저장 완료")

    def _get_hf_config(self):
        return {
            "academy_name": self.cfg.get("academy_name", ""),
            "exam_name": self.cfg.get("exam_name", ""),
            "contact_info": self.cfg.get("contact_info", ""),
            "copyright_text": self.cfg.get("copyright_text", ""),
            "show_page_number": self.cfg.get("show_page_number", False),
            "show_date": self.cfg.get("show_date", False),
        }

    def _save_to_db(self):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        if not raw:
            return messagebox.showwarning("경고", "문제 텍스트를 먼저 입력하세요.")
        parsed = parse_exam_text(raw)
        if not parsed:
            return messagebox.showerror("오류", "파싱된 문제가 없습니다.")

        title = self.ent_title.get().strip() or "untitled"
        tag_str = simpledialog.askstring("태그", "저장 태그 (쉼표 구분):",
                                         initialvalue="", parent=self.root)
        if tag_str is None:
            return

        tags = [t.strip() for t in tag_str.split(",") if t.strip()]
        db_root = self._get_db_root()
        items = exam_db.parsed_to_db_items(parsed, source_title=title, tags=tags)
        for item in items:
            exam_db.save_item(db_root, item)
            exam_db.update_index_entry(db_root, item)

        q_total = sum(len(it.get("questions", [])) for it in items)
        messagebox.showinfo("DB 저장",
            f"{len(items)}개 아이템 ({q_total}문항) 저장 완료\nDB: {db_root}")
        self._set_status(f"DB 저장: {q_total}문항 → {db_root}")

    def _open_qbank_window(self):
        db_root = self._get_db_root()
        exam_db.ensure_db_dirs(db_root)

        def on_generate_from_db(items):
            exam_data = exam_db.db_items_to_exam_data(items)
            ans_text = exam_db.db_items_to_answer_text(items)
            self.txt_exam.delete("1.0", tk.END)
            lines = []
            for q in exam_data:
                body = f"{q['num']}. {q['text']}"
                if q.get("jesi"):
                    body += f"\n<제시문>{q['jesi']}</제시문>"
                if q.get("graph_tag"):
                    body += f"\n{q['graph_tag']}"
                if q.get("bogi"):
                    body += f"\n<보기>\n{q['bogi']}"
                if q.get("choices"):
                    for j, ch in enumerate(q["choices"]):
                        if ch:
                            body += f"\n{CIRCLE_NUMS[j]} {ch}"
                lines.append(body)
            self.txt_exam.insert("1.0", "\n\n".join(lines))
            self.txt_ans.delete("1.0", tk.END)
            self.txt_ans.insert("1.0", ans_text)
            self.tabs.select(self.tab_main)
            self._set_status(f"DB에서 {len(exam_data)}문항 불러옴")

        QuestionBankWindow(self.root, db_root, on_generate_callback=on_generate_from_db)

    def _backup_db(self):
        try:
            path = exam_db.create_backup(self._get_db_root(), CONFIG_FILE)
            messagebox.showinfo("백업 완료", f"백업 저장: {path}")
            self._set_status(f"DB 백업 완료: {path}")
        except Exception as e:
            messagebox.showerror("백업 오류", str(e))

    def _restore_db(self):
        db_root = self._get_db_root()
        zip_path = filedialog.askopenfilename(
            title="백업 zip 선택", filetypes=[("Zip", "*.zip")],
            initialdir=os.path.join(db_root, "backups"))
        if not zip_path:
            return
        if not messagebox.askyesno("확인",
            f"복원하시겠습니까?\n기존 데이터는 안전 복사됩니다.\n\n{zip_path}"):
            return
        try:
            safety = exam_db.restore_backup(db_root, zip_path)
            exam_db.build_index(db_root)
            messagebox.showinfo("복원 완료", f"복원 완료!\n안전 복사: {safety}")
            self._set_status("DB 복원 완료")
        except Exception as e:
            messagebox.showerror("복원 오류", str(e))

    def _export_json(self):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        if not raw:
            return messagebox.showwarning("경고", "문제 텍스트를 먼저 입력하세요.")
        parsed = parse_exam_text(raw)
        if not parsed:
            return messagebox.showerror("오류", "파싱된 문제가 없습니다.")

        target_dir = self.dir_var.get()
        title = self.ent_title.get().strip() or "exam"
        today = datetime.datetime.now().strftime("%Y%m%d")
        try:
            path = exam_db.export_exam_json(parsed, target_dir, f"{title}_{today}",
                                            display_title=title)
            messagebox.showinfo("Export 완료", f"저장: {path}")
            self._set_status(f"JSON Export: {path}")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # ─────────────── Core: generate ───────────────
    def _get_target_dir(self, student_name=""):
        base = self.dir_var.get()
        title = self.ent_title.get().strip()
        name = student_name
        if not name and "_" in title:
            name = title.split("_")[0].strip()
        return os.path.join(base, name) if name else base

    def _generate(self, fmt):
        raw_exam = self.txt_exam.get("1.0", tk.END).strip()
        raw_ans = self.txt_ans.get("1.0", tk.END).strip()
        if not raw_exam:
            return messagebox.showwarning("경고", "문제를 입력하세요!")

        parsed = parse_exam_text(raw_exam)
        if not parsed:
            return messagebox.showerror("오류", "파싱 가능한 문제가 없습니다. 형식을 확인하세요.")

        title_raw = self.ent_title.get().strip()
        display_title = title_raw or "사탐 모의고사"
        today = datetime.datetime.now().strftime("%Y%m%d")
        filename_base = f"{title_raw}_{today}" if title_raw else f"교재_{today}"
        font_n = self.font_combo.get()
        font_s = self.size_combo.get()
        logo = self.logo_var.get()
        watermark = self.ent_watermark.get().strip()
        shuffle = self.var_shuffle.get()
        batch_raw = self.ent_batch.get().strip()
        students = [s.strip() for s in batch_raw.split(",") if s.strip()] if batch_raw else [""]
        self._save_cfg()

        generated_paths = []
        try:
            for student in students:
                target_dir = self._get_target_dir(student)
                os.makedirs(target_dir, exist_ok=True)

                data = list(parsed)
                if shuffle:
                    random.shuffle(data)

                fn = f"{student}_{filename_base}" if student else filename_base
                header_txt = f"{student} | {display_title}" if student else display_title
                footer_txt = datetime.datetime.now().strftime("%Y-%m-%d")
                hf_cfg = self._get_hf_config()

                exam_dx, exam_px = create_exam_docx(
                    target_dir, fn, data, font_n, font_s, logo,
                    display_title, watermark=watermark,
                    header=header_txt, footer=footer_txt, hf_config=hf_cfg)

                ans_dx, ans_px = None, None
                if raw_ans:
                    ans_dx, ans_px = create_answer_docx(
                        target_dir, fn, raw_ans, data, font_n, font_s,
                        logo, display_title, hf_config=hf_cfg)

                if fmt in ("pdf", "both"):
                    pairs = [(exam_dx, exam_px)]
                    if ans_dx:
                        pairs.append((ans_dx, ans_px))
                    convert_docx_pairs_to_pdf(pairs)
                    if fmt == "pdf":
                        for d in (exam_dx, ans_dx):
                            if d and os.path.exists(d):
                                os.remove(d)

                generated_paths.append(target_dir)
                add_history({
                    "title": display_title,
                    "student": student or "(단일)",
                    "q_count": len(data),
                    "format": fmt,
                    "path": target_dir,
                })

            summary = (f"총 {len(students)}명 x {len(parsed)}문항 생성 완료!"
                       if len(students) > 1
                       else f"{len(parsed)}문항 생성 완료!")
            messagebox.showinfo("완료", f"{summary}\n\n저장: {generated_paths[0]}")
            clear_draft()
            self._refresh_history()
            self._update_status()
            if generated_paths:
                self._open_dir(generated_paths[0])
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # ─────────────── Utilities ───────────────
    def _browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.dir_var.set(d)
            self._save_cfg()

    def _browse_logo(self):
        f = filedialog.askopenfilename(filetypes=[("이미지", "*.png *.jpg *.jpeg *.bmp")])
        if f:
            self.logo_var.set(f)
            self._save_cfg()

    def _open_dir(self, path):
        if not os.path.exists(path):
            path = self.dir_var.get()
        if not os.path.exists(path):
            return
        system = platform.system()
        if system == "Windows":
            os.startfile(path)
        elif system == "Darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])

    def _open_target_dir(self):
        self._open_dir(self._get_target_dir())

    def _save_cfg(self):
        self.cfg["last_dir"] = self.dir_var.get()
        self.cfg["last_logo"] = self.logo_var.get()
        self.cfg["font_name"] = self.font_combo.get()
        self.cfg["font_size"] = self.size_combo.get()
        save_config(self.cfg)

    def _set_status(self, text):
        self.lbl_status.config(text=text)

    def _update_status(self):
        """Full status update (re-parses exam text). Use sparingly."""
        raw = self.txt_exam.get("1.0", tk.END).strip()
        q_count = len(parse_exam_text(raw)) if raw else 0
        char_count = len(raw)
        self._set_status(
            f"현재 입력: {q_count}문항 ({char_count}자)  |  "
            f"문제 은행: {self._bank_count}문항  |  "
            f"히스토리: {self._hist_count}건")

    def _on_text_change(self, event=None):
        """Debounced status update — waits 500 ms after last keystroke."""
        if self._status_timer:
            self.root.after_cancel(self._status_timer)
        self._status_timer = self.root.after(500, self._update_status)

    def _save_draft_now(self, event=None):
        save_draft(
            self.txt_exam.get("1.0", tk.END).strip(),
            self.txt_ans.get("1.0", tk.END).strip(),
            self.ent_title.get().strip())
        self._set_status("임시저장 완료")

    def _start_autosave(self):
        def _tick():
            exam = self.txt_exam.get("1.0", tk.END).strip()
            if exam:
                save_draft(exam, self.txt_ans.get("1.0", tk.END).strip(),
                           self.ent_title.get().strip())
            self.root.after(30000, _tick)
        self.root.after(30000, _tick)

    def _restore_draft(self):
        draft = load_draft()
        if draft and draft.get("exam"):
            saved_at = draft.get("saved_at", "")[:19].replace("T", " ")
            if messagebox.askyesno("임시저장 복원",
                                   f"이전 임시저장({saved_at})이 있습니다.\n복원하시겠습니까?"):
                self.txt_exam.insert("1.0", draft["exam"])
                if draft.get("answer"):
                    self.txt_ans.insert("1.0", draft["answer"])
                if draft.get("title"):
                    self.ent_title.insert(0, draft["title"])
                self._update_status()

    def _on_file_drop(self, event):
        """Drag & Drop: insert .txt / .docx file content into exam text area."""
        files = self.root.tk.splitlist(event.data)
        if not files:
            return
        file_path = files[0]
        ext = os.path.splitext(file_path)[1].lower()
        extracted_text = ""
        try:
            if ext == ".txt":
                for enc in ("utf-8", "cp949", "euc-kr"):
                    try:
                        with open(file_path, "r", encoding=enc) as f:
                            extracted_text = f.read()
                        break
                    except UnicodeDecodeError:
                        continue
            elif ext == ".docx":
                doc = _docx_mod.Document(file_path)
                extracted_text = "\n".join(p.text for p in doc.paragraphs)
            else:
                messagebox.showwarning("DnD",
                    f"지원하지 않는 파일 형식: {ext}\n(.txt, .docx만 지원)")
                return

            if extracted_text:
                current = self.txt_exam.get("1.0", tk.END).strip()
                if current:
                    self.txt_exam.insert(tk.END, "\n\n" + extracted_text)
                else:
                    self.txt_exam.insert("1.0", extracted_text)
                self._update_status()
                messagebox.showinfo("DnD",
                    f"파일 로드 완료: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("DnD 오류", f"파일 파싱 중 에러: {e}")

    def _on_close(self):
        exam = self.txt_exam.get("1.0", tk.END).strip()
        if exam:
            save_draft(exam, self.txt_ans.get("1.0", tk.END).strip(),
                       self.ent_title.get().strip())
        self.root.destroy()
