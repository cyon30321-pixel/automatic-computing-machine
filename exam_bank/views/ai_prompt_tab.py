"""
AI 프롬프트 자동 생성 탭 v5.0
— 원본 지문 붙여넣기 → 문제 유형/개수 선택 → 프롬프트 자동 생성 → 복사
— AI(ChatGPT/Claude) 결과를 시험지 생성 탭에 붙여넣으면 깔끔 출력
"""

import tkinter as tk
from tkinter import messagebox, ttk

from exam_bank.constants import SCHOOL_LEVELS, GRADES, Theme
from exam_bank.services.ai_prompt import generate_ai_prompt, get_question_type_list


class AIPromptTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self._build()

    def _build(self):
        main = tk.Frame(self.frame, bg="#F0F4F8")
        main.pack(fill="both", expand=True)

        # ── 좌우 분할 ──
        paned = tk.PanedWindow(main, orient=tk.HORIZONTAL, sashwidth=6, bg="#CCC")
        paned.pack(fill="both", expand=True, padx=10, pady=10)

        # ═══ 왼쪽: 설정 + 지문 입력 ═══
        left_frame = tk.Frame(paned, bg="#F0F4F8")
        paned.add(left_frame, width=480)

        # 설정 영역
        settings_frame = tk.LabelFrame(left_frame, text="  🎯  출제 설정  ",
                                       font=("맑은 고딕", 10, "bold"),
                                       bg="#F0F4F8", fg="#1B4F72", padx=10, pady=8)
        settings_frame.pack(fill="x", padx=5, pady=(0, 8))

        # 학교급 / 학년
        row1 = tk.Frame(settings_frame, bg="#F0F4F8")
        row1.pack(fill="x", pady=3)
        tk.Label(row1, text="학교급:", bg="#F0F4F8", font=("맑은 고딕", 9)).pack(side="left", padx=(0, 3))
        self.combo_school = ttk.Combobox(row1, values=SCHOOL_LEVELS, width=8, state="readonly")
        self.combo_school.set("고등학교")
        self.combo_school.pack(side="left", padx=(0, 10))

        tk.Label(row1, text="학년:", bg="#F0F4F8", font=("맑은 고딕", 9)).pack(side="left", padx=(0, 3))
        self.combo_grade = ttk.Combobox(row1, values=GRADES, width=6, state="readonly")
        self.combo_grade.set("2학년")
        self.combo_grade.pack(side="left", padx=(0, 10))

        tk.Label(row1, text="문항수:", bg="#F0F4F8", font=("맑은 고딕", 9)).pack(side="left", padx=(0, 3))
        self.spin_num = ttk.Spinbox(row1, from_=3, to=30, width=4)
        self.spin_num.set(10)
        self.spin_num.pack(side="left")

        # 난이도
        row2 = tk.Frame(settings_frame, bg="#F0F4F8")
        row2.pack(fill="x", pady=3)
        tk.Label(row2, text="난이도:", bg="#F0F4F8", font=("맑은 고딕", 9)).pack(side="left", padx=(0, 3))
        self.combo_diff = ttk.Combobox(row2, values=["하", "중하", "중", "중상", "상", "최상"], width=6, state="readonly")
        self.combo_diff.set("중")
        self.combo_diff.pack(side="left")

        # 문제 유형 체크박스
        type_frame = tk.LabelFrame(settings_frame, text="  문제 유형 (복수 선택)  ",
                                   font=("맑은 고딕", 9), bg="#F0F4F8", fg="#555")
        type_frame.pack(fill="x", pady=5)

        self.type_vars = {}
        types = get_question_type_list()
        for i, qt in enumerate(types):
            var = tk.BooleanVar(value=(qt == "종합 (혼합)"))
            cb = tk.Checkbutton(type_frame, text=qt, variable=var,
                                bg="#F0F4F8", font=("맑은 고딕", 9),
                                activebackground="#F0F4F8")
            cb.grid(row=i // 3, column=i % 3, sticky="w", padx=5, pady=1)
            self.type_vars[qt] = var

        # 추가 지시사항
        tk.Label(settings_frame, text="추가 지시사항 (선택):",
                 bg="#F0F4F8", font=("맑은 고딕", 9), fg="#555").pack(anchor="w", pady=(5, 0))
        self.text_extra = tk.Text(settings_frame, height=3, font=("맑은 고딕", 9),
                                  relief="solid", bd=1, wrap=tk.WORD)
        self.text_extra.pack(fill="x", pady=3)

        # 지문 입력
        tk.Label(left_frame, text="  📖  원본 영어 지문을 붙여넣으세요",
                 font=("맑은 고딕", 11, "bold"), fg="#1B4F72", bg="#F0F4F8").pack(anchor="w", padx=5, pady=(5, 0))

        passage_frame = tk.Frame(left_frame, bg="#F0F4F8")
        passage_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.text_passage = tk.Text(passage_frame, wrap=tk.WORD, font=("Consolas", 10),
                                    relief="solid", bd=1, padx=8, pady=8, bg="#FFFFFF",
                                    insertbackground="#1B4F72", selectbackground="#D4E6F1")
        scroll_p = ttk.Scrollbar(passage_frame, orient="vertical", command=self.text_passage.yview)
        self.text_passage.configure(yscrollcommand=scroll_p.set)
        self.text_passage.pack(side="left", fill="both", expand=True)
        scroll_p.pack(side="right", fill="y")

        # 생성 버튼
        btn_frame = tk.Frame(left_frame, bg="#F0F4F8")
        btn_frame.pack(fill="x", padx=5, pady=8)
        tk.Button(btn_frame, text="🚀  AI 프롬프트 자동 생성",
                  font=("맑은 고딕", 12, "bold"), bg="#2E86C1", fg="white",
                  relief="flat", cursor="hand2", bd=0, padx=20, pady=8,
                  command=self._generate_prompt).pack(fill="x")

        # ═══ 오른쪽: 생성된 프롬프트 ═══
        right_frame = tk.Frame(paned, bg="#F0F4F8")
        paned.add(right_frame, width=500)

        header_right = tk.Frame(right_frame, bg="#F0F4F8")
        header_right.pack(fill="x", padx=5, pady=(0, 5))

        tk.Label(header_right, text="  📋  생성된 AI 프롬프트  (복사해서 AI에 붙여넣기)",
                 font=("맑은 고딕", 11, "bold"), fg="#1B4F72", bg="#F0F4F8").pack(side="left")

        tk.Button(header_right, text="📋 전체 복사",
                  font=("맑은 고딕", 9, "bold"), bg="#E74C3C", fg="white",
                  relief="flat", cursor="hand2", bd=0, padx=10, pady=3,
                  command=self._copy_prompt).pack(side="right", padx=5)

        prompt_frame = tk.Frame(right_frame, bg="#F0F4F8")
        prompt_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.text_prompt = tk.Text(prompt_frame, wrap=tk.WORD, font=("Consolas", 9),
                                   relief="solid", bd=1, padx=8, pady=8, bg="#FFFEF5",
                                   insertbackground="#CC8800", selectbackground="#FFF3CD")
        scroll_pr = ttk.Scrollbar(prompt_frame, orient="vertical", command=self.text_prompt.yview)
        self.text_prompt.configure(yscrollcommand=scroll_pr.set)
        self.text_prompt.pack(side="left", fill="both", expand=True)
        scroll_pr.pack(side="right", fill="y")

        # 하단 안내
        tip = tk.Label(right_frame,
                       text="💡 사용법:  프롬프트 복사 → ChatGPT/Claude에 붙여넣기 → AI 결과를 '시험지 생성' 탭에 붙여넣기 → Word/PDF 생성",
                       font=("맑은 고딕", 8), fg="#888", bg="#F0F4F8", wraplength=480)
        tip.pack(padx=5, pady=5)

    def _generate_prompt(self):
        passage = self.text_passage.get("1.0", tk.END).strip()
        if not passage:
            messagebox.showwarning("경고", "원본 영어 지문을 입력해 주세요!")
            return

        selected_types = [qt for qt, var in self.type_vars.items() if var.get()]
        if not selected_types:
            messagebox.showwarning("경고", "문제 유형을 최소 1개 이상 선택하세요!")
            return

        try:
            num_q = int(self.spin_num.get())
        except ValueError:
            num_q = 10

        prompt = generate_ai_prompt(
            passage_text=passage,
            question_types=selected_types,
            num_questions=num_q,
            difficulty=self.combo_diff.get(),
            school_level=self.combo_school.get(),
            grade=self.combo_grade.get(),
            extra_instructions=self.text_extra.get("1.0", tk.END).strip(),
        )

        self.text_prompt.delete("1.0", tk.END)
        self.text_prompt.insert("1.0", prompt)
        messagebox.showinfo("완료", f"✅ AI 프롬프트 생성 완료!\n\n오른쪽 텍스트를 복사하여 ChatGPT/Claude에 붙여넣으세요.")

    def _copy_prompt(self):
        prompt = self.text_prompt.get("1.0", tk.END).strip()
        if not prompt:
            messagebox.showwarning("알림", "먼저 프롬프트를 생성하세요!")
            return
        self.app.root.clipboard_clear()
        self.app.root.clipboard_append(prompt)
        messagebox.showinfo("복사 완료", "✅ 프롬프트가 클립보드에 복사되었습니다!\n\nChatGPT 또는 Claude에 붙여넣으세요.")
