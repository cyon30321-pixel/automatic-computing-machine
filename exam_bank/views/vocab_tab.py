"""
단어장 관리 탭 v7.0
— 단어장/단원/단어 CRUD
— AI 추출 가져오기 (JSON: 단어+예문)
— Day 조합 출제 범위 설정
— 단어 시험지 생성
— v7.0: Master-Detail 예문 UI, 예문 CRUD
"""

import os
import re
import json
import tkinter as tk
from tkinter import messagebox, filedialog, ttk

from exam_bank.constants import (
    VOCAB_CATEGORIES, SENTENCE_DIFFICULTY, SENTENCE_DIFF_SHORT,
    SENTENCE_SOURCES, AI_PROMPT_SENTENCES, AI_PROMPT_WORDS_ONLY,
)
from exam_bank.models.vocabulary import (
    add_vocab_book, update_vocab_book, delete_vocab_book, list_vocab_books, get_vocab_book,
    add_vocab_unit, update_vocab_unit, delete_vocab_unit, list_vocab_units,
    add_vocab_word, update_vocab_word, delete_vocab_word, list_vocab_words,
    get_words_by_unit_ids, bulk_add_words, bulk_add_words_with_days,
    import_words_from_excel, get_vocab_unit_stats, get_wrong_words,
    # v7.0: 예문 관련
    add_sentence, update_sentence, delete_sentence, list_sentences,
    bulk_add_sentences, get_sentences_by_word_ids,
)
from exam_bank.models.student import list_students, get_student
from exam_bank.services.vocab_generator import create_vocab_test, create_sentence_test
from exam_bank.config import open_directory


class VocabTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self.selected_units = []  # 출제 범위 (unit_id 리스트)
        self.banner_path = ""
        self._build()

    def _build(self):
        # 상단 3분할: 좌=단어장관리 | 중=출제범위 | 우=출제설정
        main_pw = tk.PanedWindow(self.frame, orient="horizontal", sashwidth=4)
        main_pw.pack(fill="both", expand=True, padx=8, pady=8)

        # ── 좌측: 단어장/단원/단어 관리 ──
        left_f = tk.Frame(main_pw, width=420)
        main_pw.add(left_f, minsize=350)
        self._build_left(left_f)

        # ── 중앙: 출제 범위 설정 ──
        mid_f = tk.Frame(main_pw, width=300)
        main_pw.add(mid_f, minsize=250)
        self._build_mid(mid_f)

        # ── 우측: 출제 설정 + 생성 ──
        right_f = tk.Frame(main_pw, width=300)
        main_pw.add(right_f, minsize=250)
        self._build_right(right_f)

    # ═══════════════════════════════════════════
    # 좌측: 단어장 관리
    # ═══════════════════════════════════════════

    def _build_left(self, parent):
        # ── 단어장 선택 ──
        bf = tk.LabelFrame(parent, text=" 📖 단어장 선택 ", font=("맑은 고딕", 9, "bold"))
        bf.pack(fill="x", padx=4, pady=4)

        r1 = tk.Frame(bf); r1.pack(fill="x", padx=6, pady=4)
        tk.Label(r1, text="단어장:").pack(side="left")
        self.combo_book = ttk.Combobox(r1, values=[], width=20, state="readonly")
        self.combo_book.pack(side="left", padx=4)
        self.combo_book.bind("<<ComboboxSelected>>", lambda e: self._on_book_selected())

        r2 = tk.Frame(bf); r2.pack(fill="x", padx=6, pady=2)
        tk.Button(r2, text="새 단어장", command=self._add_book, bg="#2563eb", fg="white", font=("맑은 고딕", 8, "bold")).pack(side="left", padx=2)
        tk.Button(r2, text="이름 수정", command=self._edit_book, font=("맑은 고딕", 8)).pack(side="left", padx=2)
        tk.Button(r2, text="삭제", command=self._delete_book, bg="#dc2626", fg="white", font=("맑은 고딕", 8)).pack(side="left", padx=2)
        tk.Button(r2, text="📋 AI 추출 가져오기", command=self._open_ai_import,
                  bg="#7c3aed", fg="white", font=("맑은 고딕", 8, "bold")).pack(side="right", padx=2)
        tk.Button(r2, text="📁 파일 가져오기", command=self._import_file,
                  font=("맑은 고딕", 8)).pack(side="right", padx=2)

        # ── 단원(Day) 목록 ──
        uf = tk.LabelFrame(parent, text=" 📋 단원 (Day) ", font=("맑은 고딕", 9, "bold"))
        uf.pack(fill="x", padx=4, pady=4)

        self.tree_units = ttk.Treeview(uf, columns=("이름", "단어수"), show="headings", height=5, selectmode="extended")
        self.tree_units.heading("이름", text="단원명")
        self.tree_units.heading("단어수", text="단어수")
        self.tree_units.column("이름", width=160)
        self.tree_units.column("단어수", width=60, anchor="center")
        self.tree_units.pack(fill="x", padx=6, pady=4)
        self.tree_units.bind("<<TreeviewSelect>>", lambda e: self._on_unit_selected())

        ur = tk.Frame(uf); ur.pack(fill="x", padx=6, pady=2)
        tk.Button(ur, text="단원 추가", command=self._add_unit, font=("맑은 고딕", 8)).pack(side="left", padx=2)
        tk.Button(ur, text="수정", command=self._edit_unit, font=("맑은 고딕", 8)).pack(side="left", padx=2)
        tk.Button(ur, text="삭제", command=self._delete_unit, bg="#dc2626", fg="white", font=("맑은 고딕", 8)).pack(side="left", padx=2)
        tk.Button(ur, text="▶ 출제 범위에 추가", command=self._add_to_scope,
                  bg="#16a34a", fg="white", font=("맑은 고딕", 8, "bold")).pack(side="right", padx=2)

        # ── 단어 목록 ──
        wf = tk.LabelFrame(parent, text=" 📝 단어 목록 ", font=("맑은 고딕", 9, "bold"))
        wf.pack(fill="both", expand=True, padx=4, pady=4)

        self.tree_words = ttk.Treeview(wf, columns=("No", "영어", "한국어", "품사"), show="headings", height=8)
        self.tree_words.heading("No", text="#")
        self.tree_words.heading("영어", text="English")
        self.tree_words.heading("한국어", text="한국어")
        self.tree_words.heading("품사", text="품사")
        self.tree_words.column("No", width=35, anchor="center")
        self.tree_words.column("영어", width=130)
        self.tree_words.column("한국어", width=130)
        self.tree_words.column("품사", width=45, anchor="center")
        scr = ttk.Scrollbar(wf, orient="vertical", command=self.tree_words.yview)
        self.tree_words.configure(yscrollcommand=scr.set)
        self.tree_words.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=4)
        scr.pack(side="right", fill="y", pady=4, padx=(0, 6))

        wr = tk.Frame(wf); wr.pack(side="bottom", fill="x", padx=6, pady=2)
        self.lbl_word_count = tk.Label(wr, text="0개", font=("맑은 고딕", 8), fg="#475569")
        self.lbl_word_count.pack(side="left")
        tk.Button(wr, text="단어 추가", command=self._add_word, font=("맑은 고딕", 8)).pack(side="right", padx=2)
        tk.Button(wr, text="수정", command=self._edit_word, font=("맑은 고딕", 8)).pack(side="right", padx=2)
        tk.Button(wr, text="삭제", command=self._delete_word, bg="#dc2626", fg="white", font=("맑은 고딕", 8)).pack(side="right", padx=2)

        # v7.0: 단어 선택 시 예문 패널 연동
        self.tree_words.bind("<<TreeviewSelect>>", lambda e: self._on_word_selected())

        # ── v7.0 예문(Sentence) Detail 패널 ──
        sf = tk.LabelFrame(parent, text=" 💬 예문 (선택된 단어) ", font=("맑은 고딕", 9, "bold"))
        sf.pack(fill="x", padx=4, pady=4)

        self.tree_sentences = ttk.Treeview(sf, columns=("난이도", "영어 예문", "한국어 해석", "target"),
                                            show="headings", height=4)
        self.tree_sentences.heading("난이도", text="난이도")
        self.tree_sentences.heading("영어 예문", text="English Sentence")
        self.tree_sentences.heading("한국어 해석", text="한국어 해석")
        self.tree_sentences.heading("target", text="target")
        self.tree_sentences.column("난이도", width=50, anchor="center")
        self.tree_sentences.column("영어 예문", width=180)
        self.tree_sentences.column("한국어 해석", width=130)
        self.tree_sentences.column("target", width=60, anchor="center")
        scr_s = ttk.Scrollbar(sf, orient="vertical", command=self.tree_sentences.yview)
        self.tree_sentences.configure(yscrollcommand=scr_s.set)
        self.tree_sentences.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=4)
        scr_s.pack(side="right", fill="y", pady=4, padx=(0, 6))

        sr = tk.Frame(sf); sr.pack(side="bottom", fill="x", padx=6, pady=2)
        self.lbl_sent_count = tk.Label(sr, text="예문 0개", font=("맑은 고딕", 8), fg="#475569")
        self.lbl_sent_count.pack(side="left")
        tk.Button(sr, text="예문 추가", command=self._add_sentence_dialog, bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 8, "bold")).pack(side="right", padx=2)
        tk.Button(sr, text="수정", command=self._edit_sentence_dialog, font=("맑은 고딕", 8)).pack(side="right", padx=2)
        tk.Button(sr, text="삭제", command=self._delete_sentence, bg="#dc2626", fg="white",
                  font=("맑은 고딕", 8)).pack(side="right", padx=2)

    # ═══════════════════════════════════════════
    # 중앙: 출제 범위
    # ═══════════════════════════════════════════

    def _build_mid(self, parent):
        sf = tk.LabelFrame(parent, text=" 🎯 출제 범위 ", font=("맑은 고딕", 9, "bold"))
        sf.pack(fill="both", expand=True, padx=4, pady=4)

        tk.Label(sf, text="선택된 단원 (여러 Day 조합 가능)", font=("맑은 고딕", 8), fg="#475569").pack(padx=6, pady=2)

        self.list_scope = tk.Listbox(sf, height=12, selectmode=tk.EXTENDED, font=("맑은 고딕", 9))
        self.list_scope.pack(fill="both", expand=True, padx=6, pady=4)

        btn_f = tk.Frame(sf); btn_f.pack(fill="x", padx=6, pady=4)
        tk.Button(btn_f, text="선택 제거", command=self._remove_from_scope, font=("맑은 고딕", 8)).pack(side="left", padx=2)
        tk.Button(btn_f, text="전체 비우기", command=self._clear_scope, font=("맑은 고딕", 8)).pack(side="left", padx=2)

        self.lbl_scope_info = tk.Label(sf, text="0개 단원 | 0개 단어", font=("맑은 고딕", 8, "bold"), fg="#2563eb")
        self.lbl_scope_info.pack(padx=6, pady=4)

        # ── 학생별 학습 이력 ──
        hf = tk.LabelFrame(parent, text=" 📊 학습 이력 ", font=("맑은 고딕", 9, "bold"))
        hf.pack(fill="x", padx=4, pady=4)

        self.tree_history = ttk.Treeview(hf, columns=("단원", "단어수", "시험횟수", "최고점", "최근"),
                                         show="headings", height=4)
        for col, w in [("단원", 100), ("단어수", 50), ("시험횟수", 55), ("최고점", 50), ("최근", 80)]:
            self.tree_history.heading(col, text=col)
            self.tree_history.column(col, width=w, anchor="center")
        self.tree_history.pack(fill="x", padx=6, pady=4)

    # ═══════════════════════════════════════════
    # 우측: 출제 설정 + 생성
    # ═══════════════════════════════════════════

    def _build_right(self, parent):
        gf = tk.LabelFrame(parent, text=" 📄 시험지 생성 ", font=("맑은 고딕", 9, "bold"))
        gf.pack(fill="both", expand=True, padx=4, pady=4)

        # 학생 선택
        r1 = tk.Frame(gf); r1.pack(fill="x", padx=8, pady=4)
        tk.Label(r1, text="학생:").pack(side="left")
        self.combo_vstudent = ttk.Combobox(r1, values=[], width=16, state="readonly")
        self.combo_vstudent.pack(side="left", padx=4)
        self.combo_vstudent.bind("<<ComboboxSelected>>", lambda e: self._refresh_history())

        # 출제 모드
        r2 = tk.Frame(gf); r2.pack(fill="x", padx=8, pady=4)
        tk.Label(r2, text="모드:").pack(side="left")
        self.combo_mode = ttk.Combobox(r2, values=["영어→한국어", "한국어→영어"], width=14, state="readonly")
        self.combo_mode.current(0); self.combo_mode.pack(side="left", padx=4)

        # 문항수
        r3 = tk.Frame(gf); r3.pack(fill="x", padx=8, pady=4)
        tk.Label(r3, text="문항수:").pack(side="left")
        self.ent_vcount = tk.Entry(r3, width=6)
        self.ent_vcount.insert(0, "40")
        self.ent_vcount.pack(side="left", padx=4)
        tk.Label(r3, text="(0 = 전체)", fg="#94a3b8", font=("맑은 고딕", 8)).pack(side="left")

        # 옵션
        r4 = tk.Frame(gf); r4.pack(fill="x", padx=8, pady=4)
        self.var_shuffle = tk.IntVar(value=1)
        tk.Checkbutton(r4, text="셔플", variable=self.var_shuffle).pack(side="left")
        self.var_hint = tk.IntVar(value=0)
        tk.Checkbutton(r4, text="첫글자 힌트", variable=self.var_hint).pack(side="left", padx=8)

        # 포맷
        r5 = tk.Frame(gf); r5.pack(fill="x", padx=8, pady=4)
        tk.Label(r5, text="포맷:").pack(side="left")
        self.combo_vformat = ttk.Combobox(r5, values=["Word + PDF", "Word만", "PDF만"], width=12, state="readonly")
        self.combo_vformat.current(0); self.combo_vformat.pack(side="left", padx=4)

        # ── 고급 설정 ──
        adv = tk.LabelFrame(gf, text=" ⚙️ 고급 설정 ", font=("맑은 고딕", 8, "bold"), fg="#475569")
        adv.pack(fill="x", padx=8, pady=6)

        tf = tk.Frame(adv); tf.pack(fill="x", padx=4, pady=2)
        tk.Label(tf, text="제목:", font=("맑은 고딕", 8)).pack(side="left")
        self.ent_vtitle = tk.Entry(tf, width=18, font=("맑은 고딕", 8))
        self.ent_vtitle.pack(side="left", padx=2)

        ff = tk.Frame(adv); ff.pack(fill="x", padx=4, pady=2)
        tk.Label(ff, text="파일명:", font=("맑은 고딕", 8)).pack(side="left")
        self.ent_vfilename = tk.Entry(ff, width=18, font=("맑은 고딕", 8))
        self.ent_vfilename.pack(side="left", padx=2)

        bnf = tk.Frame(adv); bnf.pack(fill="x", padx=4, pady=2)
        tk.Button(bnf, text="배너", font=("맑은 고딕", 8), command=self._select_banner).pack(side="left")
        tk.Button(bnf, text="✕", font=("맑은 고딕", 8), width=2, command=self._remove_banner).pack(side="left", padx=2)
        self.lbl_vbanner = tk.Label(bnf, text="(없음)", fg="#94a3b8", font=("맑은 고딕", 7))
        self.lbl_vbanner.pack(side="left", padx=2)

        tk.Label(adv, text="(비워두면 기본값 자동 적용)", font=("맑은 고딕", 7), fg="#94a3b8").pack(padx=4, pady=1)

        # ── 생성 버튼 ──
        tk.Button(gf, text="📝 단어 시험지 생성", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 11, "bold"), height=2,
                  command=self._generate_vocab_test).pack(fill="x", padx=8, pady=8)

        # 오답 단어 재시험
        tk.Button(gf, text="🔄 오답 단어만 재시험", bg="#ea580c", fg="white",
                  font=("맑은 고딕", 9, "bold"),
                  command=self._generate_wrong_only).pack(fill="x", padx=8, pady=2)

        # ── v7.0 예문 시험지 ──
        sf = tk.LabelFrame(gf, text=" 💬 예문 시험지 (v7.0) ", font=("맑은 고딕", 8, "bold"), fg="#7c3aed")
        sf.pack(fill="x", padx=8, pady=6)

        sr1 = tk.Frame(sf); sr1.pack(fill="x", padx=4, pady=2)
        tk.Label(sr1, text="모드:", font=("맑은 고딕", 8)).pack(side="left")
        self.combo_sent_mode = ttk.Combobox(sr1, values=["문장해석 (영→한)", "빈칸영작 (Cloze)"],
                                             width=16, state="readonly")
        self.combo_sent_mode.current(0); self.combo_sent_mode.pack(side="left", padx=4)

        sr2 = tk.Frame(sf); sr2.pack(fill="x", padx=4, pady=2)
        self.var_sent_hint = tk.IntVar(value=0)
        tk.Checkbutton(sr2, text="첫글자 힌트", variable=self.var_sent_hint, font=("맑은 고딕", 8)).pack(side="left")
        tk.Label(sr2, text="난이도:", font=("맑은 고딕", 8)).pack(side="left", padx=(8, 0))
        self.combo_sent_diff = ttk.Combobox(sr2, values=["전체", "1-기본", "2-핵심", "3-심화"],
                                             width=8, state="readonly")
        self.combo_sent_diff.current(0); self.combo_sent_diff.pack(side="left", padx=4)

        tk.Button(sf, text="📝 예문 시험지 생성", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 9, "bold"),
                  command=self._generate_sentence_test).pack(fill="x", padx=4, pady=4)

    # ═══════════════════════════════════════════
    # 단어장 CRUD
    # ═══════════════════════════════════════════

    def _get_current_book_id(self):
        text = self.combo_book.get()
        m = re.search(r"\[(\d+)\]", text)
        return int(m.group(1)) if m else None

    def _get_current_unit_id(self):
        sel = self.tree_units.selection()
        if not sel:
            return None
        return int(sel[0])

    def _add_book(self):
        top = tk.Toplevel(self.app.root); top.title("새 단어장"); top.geometry("350x200"); top.attributes("-topmost", True)
        tk.Label(top, text="단어장 이름:").pack(pady=(12, 2), padx=16, anchor="w")
        ent_name = tk.Entry(top, width=30); ent_name.pack(padx=16, fill="x"); ent_name.focus_set()
        tk.Label(top, text="카테고리:").pack(pady=(8, 2), padx=16, anchor="w")
        cat_f = tk.Frame(top); cat_f.pack(fill="x", padx=16)
        combo_cat = ttk.Combobox(cat_f, values=VOCAB_CATEGORIES, width=15)
        combo_cat.current(0); combo_cat.pack(side="left")
        tk.Label(cat_f, text="(직접 입력 가능)", fg="#94a3b8", font=("맑은 고딕", 8)).pack(side="left", padx=4)

        def do_add():
            name = ent_name.get().strip()
            if not name:
                return messagebox.showwarning("오류", "이름을 입력하세요.")
            add_vocab_book(self.cfg, name, combo_cat.get())
            top.destroy()
            self.refresh()

        tk.Button(top, text="생성", bg="#2563eb", fg="white", font=("맑은 고딕", 10, "bold"), command=do_add).pack(fill="x", padx=16, pady=12)
        ent_name.bind("<Return>", lambda e: do_add())

    def _edit_book(self):
        bid = self._get_current_book_id()
        if not bid:
            return messagebox.showwarning("알림", "단어장을 선택하세요.")
        book = get_vocab_book(self.cfg, bid)
        if not book:
            return

        top = tk.Toplevel(self.app.root); top.title("단어장 수정"); top.geometry("350x200"); top.attributes("-topmost", True)
        tk.Label(top, text="이름:").pack(pady=(12, 2), padx=16, anchor="w")
        ent_name = tk.Entry(top, width=30); ent_name.pack(padx=16, fill="x"); ent_name.insert(0, book["name"])
        tk.Label(top, text="카테고리:").pack(pady=(8, 2), padx=16, anchor="w")
        cat_f = tk.Frame(top); cat_f.pack(fill="x", padx=16)
        combo_cat = ttk.Combobox(cat_f, values=VOCAB_CATEGORIES, width=15)
        combo_cat.set(book.get("category", "기본단어장")); combo_cat.pack(side="left")
        tk.Label(cat_f, text="(직접 입력 가능)", fg="#94a3b8", font=("맑은 고딕", 8)).pack(side="left", padx=4)

        def do_save():
            update_vocab_book(self.cfg, bid, name=ent_name.get().strip(), category=combo_cat.get())
            top.destroy(); self.refresh()

        tk.Button(top, text="저장", bg="#16a34a", fg="white", font=("맑은 고딕", 10, "bold"), command=do_save).pack(fill="x", padx=16, pady=12)

    def _delete_book(self):
        bid = self._get_current_book_id()
        if not bid:
            return messagebox.showwarning("알림", "단어장을 선택하세요.")
        if not messagebox.askyesno("경고", "이 단어장을 삭제하시겠습니까?\n(학습 이력은 보존됩니다)"):
            return
        delete_vocab_book(self.cfg, bid)  # Soft Delete
        self.refresh()

    def _add_unit(self):
        bid = self._get_current_book_id()
        if not bid:
            return messagebox.showwarning("알림", "먼저 단어장을 선택하세요.")
        name = self._simple_input("단원 추가", "단원명 (예: Day 01):")
        if name:
            units = list_vocab_units(self.cfg, bid)
            add_vocab_unit(self.cfg, bid, name, sort_order=len(units) + 1)
            self._refresh_units()

    def _edit_unit(self):
        uid = self._get_current_unit_id()
        if not uid:
            return messagebox.showwarning("알림", "단원을 선택하세요.")
        name = self._simple_input("단원 수정", "새 이름:")
        if name:
            update_vocab_unit(self.cfg, uid, unit_name=name)
            self._refresh_units()

    def _delete_unit(self):
        uid = self._get_current_unit_id()
        if not uid:
            return messagebox.showwarning("알림", "단원을 선택하세요.")
        if not messagebox.askyesno("경고", "이 단원을 삭제하시겠습니까?"):
            return
        delete_vocab_unit(self.cfg, uid)
        self._refresh_units()
        self._refresh_words()

    def _add_word(self):
        uid = self._get_current_unit_id()
        if not uid:
            return messagebox.showwarning("알림", "단원을 선택하세요.")
        top = tk.Toplevel(self.app.root); top.title("단어 추가"); top.geometry("400x250"); top.attributes("-topmost", True)
        entries = {}
        for label_text in ["영어:", "한국어:", "품사:"]:
            f = tk.Frame(top); f.pack(fill="x", padx=16, pady=4)
            tk.Label(f, text=label_text, width=8, anchor="w").pack(side="left")
            e = tk.Entry(f, width=30); e.pack(side="left", padx=4)
            entries[label_text] = e
        entries["영어:"].focus_set()

        def do_add():
            eng = entries["영어:"].get().strip()
            kor = entries["한국어:"].get().strip()
            pos = entries["품사:"].get().strip()
            if not eng:
                return messagebox.showwarning("오류", "영어 단어를 입력하세요.")
            add_vocab_word(self.cfg, uid, eng, kor, pos)
            top.destroy()
            self._refresh_words()

        tk.Button(top, text="추가", bg="#2563eb", fg="white", font=("맑은 고딕", 10, "bold"), command=do_add).pack(fill="x", padx=16, pady=12)

    def _edit_word(self):
        sel = self.tree_words.selection()
        if not sel:
            return messagebox.showwarning("알림", "단어를 선택하세요.")
        wid = int(sel[0])
        vals = self.tree_words.item(sel[0], "values")

        top = tk.Toplevel(self.app.root); top.title("단어 수정"); top.geometry("400x250"); top.attributes("-topmost", True)
        entries = {}
        defaults = {"영어:": vals[1], "한국어:": vals[2], "품사:": vals[3]}
        for label_text in ["영어:", "한국어:", "품사:"]:
            f = tk.Frame(top); f.pack(fill="x", padx=16, pady=4)
            tk.Label(f, text=label_text, width=8, anchor="w").pack(side="left")
            e = tk.Entry(f, width=30); e.pack(side="left", padx=4)
            e.insert(0, defaults.get(label_text, ""))
            entries[label_text] = e

        def do_save():
            update_vocab_word(self.cfg, wid,
                              english=entries["영어:"].get().strip(),
                              korean=entries["한국어:"].get().strip(),
                              pos=entries["품사:"].get().strip())
            top.destroy(); self._refresh_words()

        tk.Button(top, text="저장", bg="#16a34a", fg="white", font=("맑은 고딕", 10, "bold"), command=do_save).pack(fill="x", padx=16, pady=12)

    def _delete_word(self):
        sel = self.tree_words.selection()
        if not sel:
            return messagebox.showwarning("알림", "단어를 선택하세요.")
        if not messagebox.askyesno("확인", f"{len(sel)}개 단어를 삭제하시겠습니까?"):
            return
        for s in sel:
            delete_vocab_word(self.cfg, int(s))
        self._refresh_words()

    # ═══════════════════════════════════════════
    # 출제 범위 관리
    # ═══════════════════════════════════════════

    def _add_to_scope(self):
        """선택된 단원을 출제 범위에 추가"""
        sel = self.tree_units.selection()
        if not sel:
            return messagebox.showwarning("알림", "단원을 선택하세요.")
        for s in sel:
            uid = int(s)
            if uid not in self.selected_units:
                self.selected_units.append(uid)
        self._refresh_scope()

    def _remove_from_scope(self):
        for idx in reversed(list(self.list_scope.curselection())):
            if idx < len(self.selected_units):
                self.selected_units.pop(idx)
        self._refresh_scope()

    def _clear_scope(self):
        self.selected_units.clear()
        self._refresh_scope()

    def _refresh_scope(self):
        self.list_scope.delete(0, tk.END)
        total_words = 0
        for uid in self.selected_units:
            from exam_bank.models.vocabulary import get_vocab_unit
            unit = get_vocab_unit(self.cfg, uid)
            if unit:
                book = get_vocab_book(self.cfg, unit.get("book_id"))
                book_name = book["name"] if book else "?"
                wc = unit.get("word_count", 0)
                total_words += wc
                self.list_scope.insert("end", f"{book_name} > {unit['unit_name']} ({wc}단어)")
        self.lbl_scope_info.config(text=f"{len(self.selected_units)}개 단원 | {total_words}개 단어")
        self._refresh_history()

    # ═══════════════════════════════════════════
    # AI 추출 가져오기
    # ═══════════════════════════════════════════

    def _open_ai_import(self):
        bid = self._get_current_book_id()
        if not bid:
            return messagebox.showwarning("알림", "먼저 단어장을 선택하세요.")

        top = tk.Toplevel(self.app.root)
        top.title("AI 추출 가져오기 (v7.0)")
        top.geometry("750x720")
        top.attributes("-topmost", True)

        # ── 1️⃣ 프롬프트 선택 + 복사 ──
        pf = tk.LabelFrame(top, text=" 1️⃣ AI 프롬프트 (사진과 함께 붙여넣기) ", font=("맑은 고딕", 9, "bold"))
        pf.pack(fill="x", padx=12, pady=6)

        # 프롬프트 모드 선택
        mode_f = tk.Frame(pf); mode_f.pack(fill="x", padx=8, pady=2)
        prompt_mode = tk.IntVar(value=1)  # 1=예문포함, 0=단어만
        tk.Radiobutton(mode_f, text="📝 단어 + 예문 (권장)", variable=prompt_mode, value=1,
                       font=("맑은 고딕", 8, "bold"), command=lambda: _update_prompt()).pack(side="left")
        tk.Radiobutton(mode_f, text="📋 단어만", variable=prompt_mode, value=0,
                       font=("맑은 고딕", 8), command=lambda: _update_prompt()).pack(side="left", padx=12)

        txt_prompt = tk.Text(pf, height=5, font=("Consolas", 9), bg="#f8f9fa", wrap="word")
        txt_prompt.pack(fill="x", padx=8, pady=4)

        def _get_current_prompt():
            return AI_PROMPT_SENTENCES if prompt_mode.get() == 1 else AI_PROMPT_WORDS_ONLY

        def _update_prompt():
            txt_prompt.config(state="normal")
            txt_prompt.delete("1.0", tk.END)
            txt_prompt.insert("1.0", _get_current_prompt())
            txt_prompt.config(state="disabled")

        _update_prompt()

        def copy_prompt():
            top.clipboard_clear()
            top.clipboard_append(_get_current_prompt())
            messagebox.showinfo("복사 완료", "프롬프트가 클립보드에 복사되었습니다.\nClaude/ChatGPT 채팅에 사진과 함께 붙여넣으세요.")

        tk.Button(pf, text="📋 프롬프트 복사", command=copy_prompt,
                  bg="#2563eb", fg="white", font=("맑은 고딕", 9, "bold")).pack(padx=8, pady=2)

        # ── 2️⃣ JSON 붙여넣기 ──
        jf = tk.LabelFrame(top, text=" 2️⃣ AI가 출력한 JSON 붙여넣기 ", font=("맑은 고딕", 9, "bold"))
        jf.pack(fill="x", padx=12, pady=6)
        txt_json = tk.Text(jf, height=7, font=("Consolas", 9))
        txt_json.pack(fill="x", padx=8, pady=4)

        df = tk.Frame(jf); df.pack(fill="x", padx=8, pady=4)
        tk.Label(df, text="Day 이름:").pack(side="left")
        ent_day = tk.Entry(df, width=15); ent_day.pack(side="left", padx=4)
        ent_day.insert(0, "")
        tk.Label(df, text="(비우면 section 필드로 자동 분류)", fg="#94a3b8", font=("맑은 고딕", 8)).pack(side="left")

        # ── 3️⃣ 미리보기 ──
        pvf = tk.LabelFrame(top, text=" 3️⃣ 미리보기 ", font=("맑은 고딕", 9, "bold"))
        pvf.pack(fill="both", expand=True, padx=12, pady=6)

        # v7.0: 예문 컬럼 추가
        tree_pv = ttk.Treeview(pvf, columns=("No", "영어", "한국어", "품사", "예문수"), show="headings", height=6)
        for col, w in [("No", 30), ("영어", 150), ("한국어", 150), ("품사", 50), ("예문수", 50)]:
            tree_pv.heading(col, text=col); tree_pv.column(col, width=w, anchor="center" if col in ("No", "품사", "예문수") else "w")
        tree_pv.pack(fill="both", expand=True, padx=8, pady=4)

        lbl_pv_count = tk.Label(pvf, text="", font=("맑은 고딕", 8), fg="#475569")
        lbl_pv_count.pack(padx=8)

        parsed_words = []

        def do_parse():
            nonlocal parsed_words
            raw = txt_json.get("1.0", tk.END).strip()
            if not raw:
                return messagebox.showwarning("오류", "JSON을 붙여넣으세요.")

            raw = re.sub(r'^```(?:json)?\s*', '', raw).strip()
            raw = re.sub(r'```\s*$', '', raw).strip()

            try:
                data = json.loads(raw)
            except json.JSONDecodeError as e:
                return messagebox.showerror("JSON 오류", f"JSON 파싱 실패:\n{e}\n\n앞뒤 마크다운 블록(```)이 포함되었는지 확인하세요.")

            if not isinstance(data, list):
                return messagebox.showerror("형식 오류", "JSON 배열 형식이어야 합니다.\n[{...}, {...}, ...]")

            parsed_words = data
            tree_pv.delete(*tree_pv.get_children())
            total_sents = 0
            for i, w in enumerate(data, 1):
                eng = w.get("english", w.get("word", ""))
                kor = w.get("korean", w.get("meaning", w.get("korean_meaning", "")))
                pos = w.get("pos", w.get("part_of_speech", ""))
                sents = w.get("sentences", [])
                sc = len(sents)
                total_sents += sc
                tree_pv.insert("", "end", values=(i, eng, kor, pos, sc if sc else "-"))

            # 요약 정보
            sections = {}
            for w in data:
                sec = w.get("section", "")
                if sec:
                    sections[sec] = sections.get(sec, 0) + 1

            info_parts = [f"총 {len(data)}개 단어"]
            if total_sents:
                info_parts.append(f"예문 {total_sents}개")
            if sections:
                sec_info = ", ".join(f"{k}({v})" for k, v in sections.items())
                info_parts.append(f"자동분류: {sec_info}")
            lbl_pv_count.config(text=" | ".join(info_parts))

        def do_import():
            if not parsed_words:
                return messagebox.showwarning("오류", "먼저 '미리보기'를 클릭하세요.")
            day_name = ent_day.get().strip()
            if day_name:
                results = bulk_add_words_with_days(self.cfg, bid, {day_name: parsed_words})
            else:
                has_section = any(w.get("section") for w in parsed_words)
                if has_section:
                    grouped = {}
                    for w in parsed_words:
                        sec = w.get("section", "기타").strip()
                        if sec not in grouped:
                            grouped[sec] = []
                        grouped[sec].append(w)
                    results = bulk_add_words_with_days(self.cfg, bid, grouped)
                else:
                    results = bulk_add_words_with_days(self.cfg, bid, {"Day 01": parsed_words})

            # v7.0: 예문 개수도 결과에 포함
            sent_total = results.pop("__sentences_total__", 0)
            total = sum(results.values())
            detail = "\n".join(f"  {k}: {v}개" for k, v in results.items())
            sent_msg = f"\n📝 예문 {sent_total}개 자동저장" if sent_total else ""
            messagebox.showinfo("가져오기 완료", f"총 {total}개 단어 등록!{sent_msg}\n\n{detail}")
            top.destroy()
            self.refresh()

        btn_f = tk.Frame(top); btn_f.pack(fill="x", padx=12, pady=8)
        tk.Button(btn_f, text="👁 미리보기", command=do_parse,
                  bg="#6b7280", fg="white", font=("맑은 고딕", 9, "bold")).pack(side="left", padx=4)
        tk.Button(btn_f, text="✅ DB에 저장", command=do_import,
                  bg="#16a34a", fg="white", font=("맑은 고딕", 10, "bold")).pack(side="right", padx=4)

    def _import_file(self):
        """엑셀/CSV 파일에서 단어 가져오기"""
        bid = self._get_current_book_id()
        if not bid:
            return messagebox.showwarning("알림", "먼저 단어장을 선택하세요.")
        path = filedialog.askopenfilename(
            title="단어 파일 선택",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv *.tsv")]
        )
        if not path:
            return
        try:
            results = import_words_from_excel(self.cfg, bid, path)
            total = sum(results.values())
            detail = "\n".join(f"  {k}: {v}개" for k, v in results.items())
            messagebox.showinfo("가져오기 완료", f"총 {total}개 단어 등록!\n\n{detail}")
            self.refresh()
        except Exception as e:
            messagebox.showerror("오류", f"파일 가져오기 실패:\n{e}")

    # ═══════════════════════════════════════════
    # 시험지 생성
    # ═══════════════════════════════════════════

    def _get_vstudent_id(self):
        m = re.search(r"\[(\d+)\]", self.combo_vstudent.get())
        return int(m.group(1)) if m else None

    def _generate_vocab_test(self):
        if not self.selected_units:
            return messagebox.showwarning("경고", "출제 범위를 먼저 설정하세요.\n좌측에서 단원을 선택 후 '▶ 출제 범위에 추가' 클릭!")

        # 단어 풀 가져오기
        all_words = get_words_by_unit_ids(self.cfg, self.selected_units)
        if not all_words:
            return messagebox.showwarning("경고", "선택된 범위에 단어가 없습니다.")

        # 문항수
        try:
            count = int(self.ent_vcount.get().strip() or "0")
        except ValueError:
            return messagebox.showwarning("오류", "문항수를 숫자로 입력하세요.")

        if count > 0 and count < len(all_words):
            import random
            all_words = random.sample(all_words, count)

        # 출제 모드
        exam_type = "eng_to_kor" if self.combo_mode.current() == 0 else "kor_to_eng"

        # 학생
        sid = self._get_vstudent_id()
        student_name = ""
        if sid:
            s = get_student(self.cfg, sid)
            student_name = s["name"] if s else ""

        # 단원 이름 목록
        unit_names = []
        for uid in self.selected_units:
            from exam_bank.models.vocabulary import get_vocab_unit
            u = get_vocab_unit(self.cfg, uid)
            if u:
                unit_names.append(u["unit_name"])

        messagebox.showinfo("생성 시작", "단어 시험지를 생성합니다.")

        try:
            files, err = create_vocab_test(
                target_dir=self.cfg["last_dir"],
                words=all_words,
                cfg=self.cfg,
                exam_type=exam_type,
                student_name=student_name,
                show_answer=False,
                show_first_letter=bool(self.var_hint.get()),
                shuffle=bool(self.var_shuffle.get()),
                custom_title=self.ent_vtitle.get().strip(),
                custom_filename=self.ent_vfilename.get().strip(),
                logo_path=self.banner_path or self.cfg.get("last_logo", ""),
                output_format=self.combo_vformat.get(),
                unit_names=unit_names,
            )
            if err:
                messagebox.showwarning("경고", f"생성 완료 (일부 오류):\n{err}")
            else:
                fnames = "\n".join(os.path.basename(f) for f in files)
                messagebox.showinfo("완료", f"생성 완료!\n\n{fnames}\n\n저장: {self.cfg['last_dir']}")
            open_directory(self.cfg["last_dir"])
        except PermissionError:
            messagebox.showerror("파일 접근 오류", "이전 파일이 열려 있습니다.\n먼저 닫은 후 다시 시도해주세요.")
        except Exception as e:
            messagebox.showerror("생성 오류", f"오류:\n{e}")

    def _generate_wrong_only(self):
        """오답 단어만 모아서 재시험"""
        sid = self._get_vstudent_id()
        if not sid:
            return messagebox.showwarning("알림", "학생을 선택하세요.")

        unit_ids = self.selected_units if self.selected_units else None
        wrong = get_wrong_words(self.cfg, sid, unit_ids)
        if not wrong:
            return messagebox.showinfo("결과", "틀린 단어가 없습니다!")

        exam_type = "eng_to_kor" if self.combo_mode.current() == 0 else "kor_to_eng"
        s = get_student(self.cfg, sid)
        student_name = s["name"] if s else ""

        try:
            files, err = create_vocab_test(
                target_dir=self.cfg["last_dir"],
                words=wrong,
                cfg=self.cfg,
                exam_type=exam_type,
                student_name=student_name,
                shuffle=True,
                custom_title=f"오답 재시험 [{student_name}]",
                logo_path=self.banner_path or self.cfg.get("last_logo", ""),
                output_format=self.combo_vformat.get(),
            )
            if not err:
                messagebox.showinfo("완료", f"오답 재시험지 생성!\n{len(wrong)}개 단어\n\n저장: {self.cfg['last_dir']}")
            open_directory(self.cfg["last_dir"])
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _generate_sentence_test(self):
        """v7.0: 예문 시험지 생성"""
        if not self.selected_units:
            return messagebox.showwarning("경고", "출제 범위를 먼저 설정하세요.\n좌측에서 단원을 선택 후 '▶ 출제 범위에 추가' 클릭!")

        # 출제 범위의 단어 ID 수집
        all_words = get_words_by_unit_ids(self.cfg, self.selected_units)
        if not all_words:
            return messagebox.showwarning("경고", "선택된 범위에 단어가 없습니다.")

        word_ids = [w["id"] for w in all_words]
        word_map = {w["id"]: w for w in all_words}

        # 예문 가져오기
        sentences = get_sentences_by_word_ids(self.cfg, word_ids)
        if not sentences:
            return messagebox.showwarning("경고", "선택된 범위에 예문이 없습니다.\nAI 가져오기에서 '단어 + 예문' 모드로 먼저 추가하세요.")

        # 난이도 필터
        diff_sel = self.combo_sent_diff.get()
        if diff_sel.startswith("1"):
            sentences = [s for s in sentences if s.get("difficulty", 1) == 1]
        elif diff_sel.startswith("2"):
            sentences = [s for s in sentences if s.get("difficulty", 1) == 2]
        elif diff_sel.startswith("3"):
            sentences = [s for s in sentences if s.get("difficulty", 1) == 3]

        if not sentences:
            return messagebox.showwarning("경고", f"난이도 '{diff_sel}' 예문이 없습니다.")

        # 예문 아이템 조합 (word 정보 포함)
        items = []
        for s in sentences:
            wid = s.get("word_id")
            w = word_map.get(wid, {})
            items.append({
                "sentence_en": s.get("sentence_en", ""),
                "sentence_ko": s.get("sentence_ko", ""),
                "target_form": s.get("target_form", ""),
                "difficulty": s.get("difficulty", 1),
                "word_english": w.get("english", ""),
                "word_korean": w.get("korean", ""),
            })

        # 모드
        test_mode = "translation" if self.combo_sent_mode.current() == 0 else "cloze"

        # 학생
        sid = self._get_vstudent_id()
        student_name = ""
        if sid:
            s = get_student(self.cfg, sid)
            student_name = s["name"] if s else ""

        # 단원 이름
        unit_names = []
        for uid in self.selected_units:
            from exam_bank.models.vocabulary import get_vocab_unit
            u = get_vocab_unit(self.cfg, uid)
            if u:
                unit_names.append(u["unit_name"])

        mode_label = "문장해석" if test_mode == "translation" else "빈칸영작"
        messagebox.showinfo("생성 시작", f"예문 시험지({mode_label}) {len(items)}문항 생성합니다.")

        try:
            files, err = create_sentence_test(
                target_dir=self.cfg["last_dir"],
                sentence_items=items,
                cfg=self.cfg,
                test_mode=test_mode,
                student_name=student_name,
                show_hint=bool(self.var_sent_hint.get()),
                shuffle=bool(self.var_shuffle.get()),
                custom_title=self.ent_vtitle.get().strip(),
                custom_filename=self.ent_vfilename.get().strip(),
                logo_path=self.banner_path or self.cfg.get("last_logo", ""),
                output_format=self.combo_vformat.get(),
                unit_names=unit_names,
            )
            if err:
                messagebox.showwarning("경고", f"생성 완료 (일부 오류):\n{err}")
            else:
                fnames = "\n".join(os.path.basename(f) for f in files)
                messagebox.showinfo("완료", f"예문 시험지 생성 완료!\n\n{fnames}\n\n저장: {self.cfg['last_dir']}")
            open_directory(self.cfg["last_dir"])
        except PermissionError:
            messagebox.showerror("파일 접근 오류", "이전 파일이 열려 있습니다.\n먼저 닫은 후 다시 시도해주세요.")
        except Exception as e:
            messagebox.showerror("생성 오류", f"오류:\n{e}")

    # ═══════════════════════════════════════════
    # 배너
    # ═══════════════════════════════════════════

    def _select_banner(self):
        path = filedialog.askopenfilename(
            title="배너 이미지 선택",
            filetypes=[("이미지 파일", "*.png *.jpg *.jpeg *.bmp *.gif")]
        )
        if path:
            self.banner_path = path
            self.lbl_vbanner.config(text=os.path.basename(path)[:18])

    def _remove_banner(self):
        self.banner_path = ""
        self.lbl_vbanner.config(text="(없음)")

    # ═══════════════════════════════════════════
    # 새로고침
    # ═══════════════════════════════════════════

    def refresh(self):
        """전체 새로고침"""
        # 단어장 목록
        books = list_vocab_books(self.cfg)
        self.combo_book["values"] = [f"[{b['id']}] {b['name']} ({b.get('category','')})" for b in books]
        if books and not self.combo_book.get():
            self.combo_book.current(0)

        # 학생 목록
        students = list_students(self.cfg)
        self.combo_vstudent["values"] = ["(선택안함)"] + [f"[{s['id']}] {s['name']}" for s in students]
        if self.combo_vstudent["values"] and not self.combo_vstudent.get():
            self.combo_vstudent.current(0)

        self._on_book_selected()

    def _on_book_selected(self):
        """단어장 선택 시 단원 목록 새로고침"""
        self._refresh_units()
        self._refresh_words()

    def _refresh_units(self):
        self.tree_units.delete(*self.tree_units.get_children())
        bid = self._get_current_book_id()
        if not bid:
            return
        units = list_vocab_units(self.cfg, bid)
        for u in units:
            self.tree_units.insert("", "end", iid=str(u["id"]),
                                   values=(u["unit_name"], u.get("word_count", 0)))

    def _on_unit_selected(self):
        self._refresh_words()

    def _refresh_words(self):
        self.tree_words.delete(*self.tree_words.get_children())
        # v7.0: 예문 패널도 초기화
        self.tree_sentences.delete(*self.tree_sentences.get_children())
        self.lbl_sent_count.config(text="예문 0개")

        uid = self._get_current_unit_id()
        if not uid:
            self.lbl_word_count.config(text="0개")
            return
        words = list_vocab_words(self.cfg, uid)
        for i, w in enumerate(words, 1):
            self.tree_words.insert("", "end", iid=str(w["id"]),
                                   values=(i, w["english"], w["korean"], w.get("part_of_speech", "")))
        self.lbl_word_count.config(text=f"{len(words)}개")

    def _refresh_history(self):
        """학습 이력 트리뷰 새로고침"""
        self.tree_history.delete(*self.tree_history.get_children())
        sid = self._get_vstudent_id()
        bid = self._get_current_book_id()
        if not sid or not bid:
            return
        stats = get_vocab_unit_stats(self.cfg, bid, sid)
        for s in stats:
            self.tree_history.insert("", "end", values=(
                s["unit_name"], s["word_count"], s["exam_count"],
                f"{s['best_score']}%" if s["best_score"] else "-",
                s["last_date"] or "-"
            ))

    # ═══════════════════════════════════════════
    # v7.0: 예문 (Sentence) CRUD
    # ═══════════════════════════════════════════

    def _get_selected_word_id(self):
        """단어 Treeview에서 선택된 word_id 반환"""
        sel = self.tree_words.selection()
        return int(sel[0]) if sel else None

    def _get_selected_sentence_id(self):
        """예문 Treeview에서 선택된 sentence_id 반환"""
        sel = self.tree_sentences.selection()
        return int(sel[0]) if sel else None

    def _on_word_selected(self):
        """단어 클릭 → 예문 패널 새로고침 (Master-Detail)"""
        self._refresh_sentences()

    def _refresh_sentences(self):
        """선택된 단어의 예문 목록 새로고침"""
        self.tree_sentences.delete(*self.tree_sentences.get_children())
        wid = self._get_selected_word_id()
        if not wid:
            self.lbl_sent_count.config(text="예문 0개")
            return
        sents = list_sentences(self.cfg, wid)
        for s in sents:
            diff_label = SENTENCE_DIFF_SHORT.get(s.get("difficulty", 1), "?")
            en = s.get("sentence_en", "")
            ko = s.get("sentence_ko", "")
            tgt = s.get("target_form", "")
            self.tree_sentences.insert("", "end", iid=str(s["id"]),
                                        values=(diff_label, en, ko, tgt))
        self.lbl_sent_count.config(text=f"예문 {len(sents)}개")

    def _add_sentence_dialog(self):
        """예문 수동 추가 대화상자"""
        wid = self._get_selected_word_id()
        if not wid:
            return messagebox.showwarning("알림", "먼저 단어를 선택하세요.")

        # 선택된 단어의 english 가져오기
        sel = self.tree_words.selection()
        word_vals = self.tree_words.item(sel[0], "values") if sel else None
        word_eng = word_vals[1] if word_vals else ""

        top = tk.Toplevel(self.app.root)
        top.title(f"예문 추가 — {word_eng}")
        top.geometry("550x340")
        top.attributes("-topmost", True)

        # 영어 예문
        f1 = tk.Frame(top); f1.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(f1, text="영어 예문:", width=10, anchor="w").pack(side="left")
        ent_en = tk.Entry(f1, width=45, font=("맑은 고딕", 9)); ent_en.pack(side="left", padx=4)
        ent_en.focus_set()

        # 한국어 해석
        f2 = tk.Frame(top); f2.pack(fill="x", padx=16, pady=4)
        tk.Label(f2, text="한국어 해석:", width=10, anchor="w").pack(side="left")
        ent_ko = tk.Entry(f2, width=45, font=("맑은 고딕", 9)); ent_ko.pack(side="left", padx=4)

        # target form
        f3 = tk.Frame(top); f3.pack(fill="x", padx=16, pady=4)
        tk.Label(f3, text="target form:", width=10, anchor="w").pack(side="left")
        ent_tgt = tk.Entry(f3, width=20, font=("맑은 고딕", 9)); ent_tgt.pack(side="left", padx=4)
        ent_tgt.insert(0, word_eng)
        tk.Label(f3, text="(문장 내 실제 활용형)", fg="#94a3b8", font=("맑은 고딕", 8)).pack(side="left")

        # 난이도
        f4 = tk.Frame(top); f4.pack(fill="x", padx=16, pady=4)
        tk.Label(f4, text="난이도:", width=10, anchor="w").pack(side="left")
        diff_vals = [f"{k} - {v}" for k, v in SENTENCE_DIFFICULTY.items()]
        combo_diff = ttk.Combobox(f4, values=diff_vals, width=18, state="readonly")
        combo_diff.current(0); combo_diff.pack(side="left", padx=4)

        # 출처
        f5 = tk.Frame(top); f5.pack(fill="x", padx=16, pady=4)
        tk.Label(f5, text="출처:", width=10, anchor="w").pack(side="left")
        combo_src = ttk.Combobox(f5, values=SENTENCE_SOURCES, width=18)
        combo_src.set("직접입력"); combo_src.pack(side="left", padx=4)

        def do_add():
            en = ent_en.get().strip()
            if not en:
                return messagebox.showwarning("오류", "영어 예문을 입력하세요.")
            diff = int(combo_diff.get().split(" - ")[0]) if combo_diff.get() else 1
            add_sentence(
                self.cfg, wid,
                sentence_en=en,
                sentence_ko=ent_ko.get().strip(),
                target_form=ent_tgt.get().strip(),
                difficulty=diff,
                source=combo_src.get().strip(),
            )
            top.destroy()
            self._refresh_sentences()

        tk.Button(top, text="추가", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"), command=do_add).pack(fill="x", padx=16, pady=12)
        ent_en.bind("<Return>", lambda e: ent_ko.focus_set())
        ent_ko.bind("<Return>", lambda e: do_add())

    def _edit_sentence_dialog(self):
        """예문 수정 대화상자"""
        sid = self._get_selected_sentence_id()
        if not sid:
            return messagebox.showwarning("알림", "예문을 선택하세요.")

        vals = self.tree_sentences.item(str(sid), "values")
        # vals = (난이도, 영어 예문, 한국어 해석, target)
        curr_diff_label = vals[0]
        curr_en = vals[1]
        curr_ko = vals[2]
        curr_tgt = vals[3]

        top = tk.Toplevel(self.app.root)
        top.title("예문 수정")
        top.geometry("550x320")
        top.attributes("-topmost", True)

        f1 = tk.Frame(top); f1.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(f1, text="영어 예문:", width=10, anchor="w").pack(side="left")
        ent_en = tk.Entry(f1, width=45, font=("맑은 고딕", 9)); ent_en.pack(side="left", padx=4)
        ent_en.insert(0, curr_en); ent_en.focus_set()

        f2 = tk.Frame(top); f2.pack(fill="x", padx=16, pady=4)
        tk.Label(f2, text="한국어 해석:", width=10, anchor="w").pack(side="left")
        ent_ko = tk.Entry(f2, width=45, font=("맑은 고딕", 9)); ent_ko.pack(side="left", padx=4)
        ent_ko.insert(0, curr_ko)

        f3 = tk.Frame(top); f3.pack(fill="x", padx=16, pady=4)
        tk.Label(f3, text="target form:", width=10, anchor="w").pack(side="left")
        ent_tgt = tk.Entry(f3, width=20, font=("맑은 고딕", 9)); ent_tgt.pack(side="left", padx=4)
        ent_tgt.insert(0, curr_tgt)

        f4 = tk.Frame(top); f4.pack(fill="x", padx=16, pady=4)
        tk.Label(f4, text="난이도:", width=10, anchor="w").pack(side="left")
        diff_vals = [f"{k} - {v}" for k, v in SENTENCE_DIFFICULTY.items()]
        combo_diff = ttk.Combobox(f4, values=diff_vals, width=18, state="readonly")
        # 현재 난이도 매칭
        diff_idx = 0
        for i, (k, v) in enumerate(SENTENCE_DIFF_SHORT.items()):
            if v == curr_diff_label:
                diff_idx = i; break
        combo_diff.current(diff_idx); combo_diff.pack(side="left", padx=4)

        def do_save():
            diff = int(combo_diff.get().split(" - ")[0]) if combo_diff.get() else 1
            update_sentence(
                self.cfg, sid,
                sentence_en=ent_en.get().strip(),
                sentence_ko=ent_ko.get().strip(),
                target_form=ent_tgt.get().strip(),
                difficulty=diff,
            )
            top.destroy()
            self._refresh_sentences()

        tk.Button(top, text="저장", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 10, "bold"), command=do_save).pack(fill="x", padx=16, pady=12)

    def _delete_sentence(self):
        """선택된 예문 삭제"""
        sid = self._get_selected_sentence_id()
        if not sid:
            return messagebox.showwarning("알림", "예문을 선택하세요.")
        if not messagebox.askyesno("확인", "이 예문을 삭제하시겠습니까?"):
            return
        delete_sentence(self.cfg, sid)
        self._refresh_sentences()

    # ═══════════════════════════════════════════
    # 유틸리티
    # ═══════════════════════════════════════════

    def _simple_input(self, title, prompt):
        top = tk.Toplevel(self.app.root); top.title(title); top.geometry("300x120"); top.attributes("-topmost", True)
        tk.Label(top, text=prompt).pack(pady=(12, 4), padx=16, anchor="w")
        ent = tk.Entry(top, width=30); ent.pack(padx=16, fill="x"); ent.focus_set()
        result = [None]

        def do_ok():
            val = ent.get().strip()
            if val:
                result[0] = val
            top.destroy()

        tk.Button(top, text="확인", command=do_ok, bg="#2563eb", fg="white").pack(fill="x", padx=16, pady=8)
        ent.bind("<Return>", lambda e: do_ok())
        top.wait_window()
        return result[0]
