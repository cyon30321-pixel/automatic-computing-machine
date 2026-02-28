"""
시험지 생성 탭 v5.0
— Exam Maker PRO v2.1 의 시험지/워크북/구문해석지 생성 기능 통합
"""

import os
import datetime
import tkinter as tk
from tkinter import messagebox, filedialog, ttk

from exam_bank.constants import Theme
from exam_bank.config import open_directory
from exam_bank.services.generator import (
    create_full_exam_docx, create_answer_sheet_docx,
    create_workbook_docx, create_syntax_workbook_docx,
    convert_docx_pairs_to_pdf_checked,
)


class ExamTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self._build()

    def _build(self):
        main = tk.Frame(self.frame, bg="#F0F4F8")
        main.pack(fill="both", expand=True)

        # ── 설정 프레임 ──
        frame_top = tk.LabelFrame(main, text="  ⚙️  생성 설정  ",
                                  font=("맑은 고딕", 10, "bold"),
                                  bg="#F0F4F8", fg="#1B4F72", padx=10, pady=8)
        frame_top.pack(fill="x", padx=15, pady=(10, 5))

        # 저장 폴더
        tk.Label(frame_top, text="저장 폴더:", bg="#F0F4F8", font=("맑은 고딕", 9)).grid(row=0, column=0, sticky="e", padx=5)
        self.dir_var = tk.StringVar(value=self.cfg.get("last_dir", ""))
        tk.Entry(frame_top, textvariable=self.dir_var, width=45, font=("맑은 고딕", 9)).grid(row=0, column=1, padx=5)
        tk.Button(frame_top, text="📁 찾아보기", font=("맑은 고딕", 8),
                  command=self._browse_dir).grid(row=0, column=2, padx=3)
        tk.Button(frame_top, text="📂 열기", font=("맑은 고딕", 8),
                  command=self._open_folder).grid(row=0, column=3, padx=3)

        # 로고
        tk.Label(frame_top, text="배너 로고:", bg="#F0F4F8", font=("맑은 고딕", 9)).grid(row=1, column=0, sticky="e", padx=5, pady=3)
        self.logo_var = tk.StringVar(value=self.cfg.get("last_logo", ""))
        tk.Entry(frame_top, textvariable=self.logo_var, width=45, font=("맑은 고딕", 9)).grid(row=1, column=1, padx=5)
        tk.Button(frame_top, text="🖼️ 선택", font=("맑은 고딕", 8),
                  command=self._browse_logo).grid(row=1, column=2, padx=3)

        # 파일 제목
        tk.Label(frame_top, text="파일 제목:", bg="#F0F4F8", font=("맑은 고딕", 9)).grid(row=2, column=0, sticky="e", padx=5, pady=3)
        self.entry_file = tk.Entry(frame_top, width=45, font=("맑은 고딕", 9))
        self.entry_file.grid(row=2, column=1, padx=5)
        tk.Label(frame_top, text="예: 김가은_중간고사 → '김가은' 폴더 자동생성",
                 fg="#888", bg="#F0F4F8", font=("맑은 고딕", 7)).grid(row=2, column=2, columnspan=2, sticky="w")

        # 글꼴
        tk.Label(frame_top, text="글꼴/크기:", bg="#F0F4F8", font=("맑은 고딕", 9)).grid(row=3, column=0, sticky="e", padx=5, pady=3)
        frame_font = tk.Frame(frame_top, bg="#F0F4F8")
        frame_font.grid(row=3, column=1, sticky="w", pady=3)
        self.font_combo = ttk.Combobox(frame_font, values=["맑은 고딕", "바탕", "굴림", "나눔고딕", "나눔스퀘어"], width=14)
        self.font_combo.set(self.cfg.get("font_name", "맑은 고딕"))
        self.font_combo.pack(side="left", padx=(5, 10))
        self.size_combo = ttk.Combobox(frame_font, values=["9", "10", "11", "12"], width=5)
        self.size_combo.set(self.cfg.get("font_size", "10"))
        self.size_combo.pack(side="left")
        tk.Label(frame_font, text="pt", bg="#F0F4F8", font=("맑은 고딕", 8), fg="#888").pack(side="left")

        # ── 지문/문제 입력 ──
        lbl_exam = tk.Label(main, text="  📝  지문/문제 혹은 구문을 붙여넣으세요  (AI 결과 붙여넣기 가능)",
                            font=("맑은 고딕", 11, "bold"), fg="#1B4F72", bg="#F0F4F8")
        lbl_exam.pack(anchor="w", padx=15, pady=(10, 0))

        text_frame_1 = tk.Frame(main, bg="#F0F4F8")
        text_frame_1.pack(padx=15, pady=5, fill="both", expand=True)
        self.text_exam = tk.Text(text_frame_1, wrap=tk.WORD, width=100, height=14,
                                 font=("Consolas", 10), relief="solid", bd=1,
                                 padx=8, pady=8, bg="#FFFFFF",
                                 insertbackground="#1B4F72", selectbackground="#D4E6F1")
        scroll_1 = ttk.Scrollbar(text_frame_1, orient="vertical", command=self.text_exam.yview)
        self.text_exam.configure(yscrollcommand=scroll_1.set)
        self.text_exam.pack(side="left", fill="both", expand=True)
        scroll_1.pack(side="right", fill="y")

        # ── 정답/해설 ──
        lbl_ans = tk.Label(main, text="  💡  정답 및 해설 (선택사항)",
                           font=("맑은 고딕", 11, "bold"), fg="#27AE60", bg="#F0F4F8")
        lbl_ans.pack(anchor="w", padx=15, pady=(8, 0))

        text_frame_2 = tk.Frame(main, bg="#F0F4F8")
        text_frame_2.pack(padx=15, pady=5, fill="both", expand=True)
        self.text_ans = tk.Text(text_frame_2, wrap=tk.WORD, width=100, height=6,
                                font=("Consolas", 10), relief="solid", bd=1,
                                padx=8, pady=8, bg="#FFFFFF",
                                insertbackground="#27AE60", selectbackground="#D5F5E3")
        scroll_2 = ttk.Scrollbar(text_frame_2, orient="vertical", command=self.text_ans.yview)
        self.text_ans.configure(yscrollcommand=scroll_2.set)
        self.text_ans.pack(side="left", fill="both", expand=True)
        scroll_2.pack(side="right", fill="y")

        # ── 생성 버튼 ──
        frame_buttons = tk.Frame(main, bg="#F0F4F8")
        frame_buttons.pack(pady=8)

        def make_btn(parent, text, bg_color, command, width=11):
            return tk.Button(parent, text=text, bg=bg_color, fg="white",
                             font=("맑은 고딕", 9, "bold"), command=command,
                             width=width, relief="flat", cursor="hand2",
                             activebackground=bg_color, bd=0, padx=8, pady=4)

        # 모의고사
        frame_exam = tk.LabelFrame(frame_buttons, text="  📝 모의고사 & 정답지  ",
                                   font=("맑은 고딕", 9, "bold"), fg="#1B4F72",
                                   bg="#F0F4F8", padx=5, pady=5)
        frame_exam.pack(fill="x", pady=3, padx=5)
        make_btn(frame_exam, "Word 📄", "#2E86C1", lambda: self._generate("exam", "word")).pack(side="left", padx=6, pady=3)
        make_btn(frame_exam, "PDF 📕", "#E74C3C", lambda: self._generate("exam", "pdf")).pack(side="left", padx=6, pady=3)
        make_btn(frame_exam, "Word+PDF 📦", "#1B4F72", lambda: self._generate("exam", "both"), 13).pack(side="left", padx=6, pady=3)

        # 워크북
        frame_wb = tk.LabelFrame(frame_buttons, text="  📘 본문 워크북  ",
                                 font=("맑은 고딕", 9, "bold"), fg="#6C3483",
                                 bg="#F0F4F8", padx=5, pady=5)
        frame_wb.pack(fill="x", pady=3, padx=5)
        make_btn(frame_wb, "Word 📄", "#17A589", lambda: self._generate("workbook", "word")).pack(side="left", padx=6, pady=3)
        make_btn(frame_wb, "PDF 📕", "#CB4335", lambda: self._generate("workbook", "pdf")).pack(side="left", padx=6, pady=3)
        make_btn(frame_wb, "Word+PDF 📦", "#6C3483", lambda: self._generate("workbook", "both"), 13).pack(side="left", padx=6, pady=3)

        # 구문 해석지
        frame_syn = tk.LabelFrame(frame_buttons, text="  🍀 구문 & 단어 해석지  ",
                                  font=("맑은 고딕", 9, "bold"), fg="#27AE60",
                                  bg="#F0F4F8", padx=5, pady=5)
        frame_syn.pack(fill="x", pady=3, padx=5)
        make_btn(frame_syn, "Word 📄", "#27AE60", lambda: self._generate("syntax", "word")).pack(side="left", padx=6, pady=3)
        make_btn(frame_syn, "PDF 📕", "#E67E22", lambda: self._generate("syntax", "pdf")).pack(side="left", padx=6, pady=3)
        make_btn(frame_syn, "Word+PDF 📦", "#1E8449", lambda: self._generate("syntax", "both"), 13).pack(side="left", padx=6, pady=3)

    def _get_target_dir(self):
        base_dir = self.dir_var.get() or self.cfg.get("last_dir", "")
        raw_t = self.entry_file.get().strip()
        if "_" in raw_t:
            student_name = raw_t.split("_")[0].strip()
            if student_name:
                return os.path.join(base_dir, student_name)
        return base_dir

    def _browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.dir_var.set(d)
            self.cfg["last_dir"] = d

    def _browse_logo(self):
        f = filedialog.askopenfilename(filetypes=[("이미지", "*.png *.jpg *.jpeg *.bmp")])
        if f:
            self.logo_var.set(f)
            self.cfg["last_logo"] = f

    def _open_folder(self):
        target = self._get_target_dir()
        if os.path.exists(target):
            open_directory(target)
        else:
            base = self.dir_var.get()
            if os.path.exists(base):
                open_directory(base)
            else:
                messagebox.showwarning("알림", "폴더가 존재하지 않습니다.")

    def _generate(self, doc_type, format_type):
        raw_e = self.text_exam.get("1.0", tk.END).strip()
        raw_a = self.text_ans.get("1.0", tk.END).strip()
        raw_t = self.entry_file.get().strip()

        if not raw_e:
            messagebox.showwarning("경고", "지문/문제를 입력해 주세요!")
            return

        today_str = datetime.datetime.now().strftime("%Y%m%d")
        display_t = raw_t if raw_t else "실전 모의고사"
        filename = f"{raw_t}_{today_str}" if raw_t else f"교재_{today_str}"
        font_n = self.font_combo.get()
        font_s = self.size_combo.get()
        logo_p = self.logo_var.get()
        target_dir = self._get_target_dir()
        os.makedirs(target_dir, exist_ok=True)

        try:
            self.app.root.config(cursor="wait")
            self.app.root.update()

            if doc_type == "exam":
                exam_d, exam_p = create_full_exam_docx(target_dir, filename, raw_e, font_n, font_s, logo_p, display_t)
                if raw_a:
                    ans_d, ans_p = create_answer_sheet_docx(target_dir, filename, raw_a, font_n, font_s, logo_p, display_t)
                else:
                    ans_d, ans_p = None, None
            elif doc_type == "workbook":
                exam_d, exam_p = create_workbook_docx(target_dir, filename, raw_e, font_n, font_s, logo_p, display_t)
                ans_d, ans_p = None, None
            elif doc_type == "syntax":
                exam_d, exam_p = create_syntax_workbook_docx(target_dir, filename, raw_e, font_n, font_s, logo_p, display_t)
                ans_d, ans_p = None, None
            else:
                return

            if format_type in ["pdf", "both"]:
                pairs = [(exam_d, exam_p)]
                if ans_d:
                    pairs.append((ans_d, ans_p))
                convert_docx_pairs_to_pdf_checked(pairs)
                if format_type == "pdf":
                    for d in [exam_d, ans_d]:
                        if d and os.path.exists(d):
                            os.remove(d)

            doc_name = {"exam": "모의고사", "workbook": "워크북", "syntax": "구문 해석지"}[doc_type]
            messagebox.showinfo("완료", f"✅ {doc_name} 생성 완료!\n📂 위치: {target_dir}")

        except Exception as e:
            messagebox.showerror("오류", str(e))
        finally:
            self.app.root.config(cursor="")
