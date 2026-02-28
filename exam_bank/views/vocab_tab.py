"""
단어장 관리 탭 v7.1
— 단어장/단원/단어 CRUD
— AI 추출 가져오기 (JSON: 단어+예문)
— Day 조합 + 개별 단어 출제 범위 설정
— 단어 시험지 / 예문 시험지 생성
— v7.1 개선:
  ✅ Ctrl+A / Shift+클릭 / Ctrl+클릭 다중선택
  ✅ 개별 단어 출제범위 추가
  ✅ 카테고리 필터 시스템 (기억 + 자동표시)
  ✅ 새 단어장 다이얼로그 개선
  ✅ 예문 시험지 생성 버그 수정 (dict→list 변환)
  ✅ 오답 재시험 버튼 제거
  ✅ 예문 유무 컬럼 추가
  ✅ UI 레이아웃 상단→하단 워크플로우 개선
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
        self.selected_units = []       # 출제 범위 (unit_id 리스트)
        self.selected_word_ids = []    # v7.1: 개별 단어 출제범위
        self.banner_path = ""
        self._word_sentence_cache = {} # {word_id: bool} 예문 유무 캐시
        self._build()

    def _build(self):
        # v7.1: 좌(단어장관리) | 우(출제설정+생성) 2단 구조
        main_pw = tk.PanedWindow(self.frame, orient="horizontal", sashwidth=5,
                                  bg="#e2e8f0")
        main_pw.pack(fill="both", expand=True, padx=6, pady=6)

        # ── 좌측: 단어장 관리 (넓게) ──
        left_f = tk.Frame(main_pw, width=520)
        main_pw.add(left_f, minsize=450)
        self._build_left(left_f)

        # ── 우측: 출제범위 + 설정 + 생성 ──
        right_f = tk.Frame(main_pw, width=380)
        main_pw.add(right_f, minsize=320)
        self._build_right(right_f)

    # ═══════════════════════════════════════════
    # 좌측: 단어장 관리
    # ═══════════════════════════════════════════

    def _build_left(self, parent):
        # ── 단어장 선택 + 카테고리 필터 ──
        bf = tk.LabelFrame(parent, text=" 📖 단어장 ", font=("맑은 고딕", 9, "bold"))
        bf.pack(fill="x", padx=4, pady=(4, 2))

        # v7.1: 카테고리 필터 행
        r0 = tk.Frame(bf); r0.pack(fill="x", padx=6, pady=(4, 2))
        tk.Label(r0, text="카테고리:", font=("맑은 고딕", 8)).pack(side="left")
        self.combo_category = ttk.Combobox(r0, values=["전체"] + list(VOCAB_CATEGORIES),
                                            width=12, state="readonly", font=("맑은 고딕", 8))
        self.combo_category.current(0)
        self.combo_category.pack(side="left", padx=4)
        self.combo_category.bind("<<ComboboxSelected>>", lambda e: self._on_category_changed())

        tk.Label(r0, text="단어장:", font=("맑은 고딕", 8)).pack(side="left", padx=(8, 0))
        self.combo_book = ttk.Combobox(r0, values=[], width=22, state="readonly", font=("맑은 고딕", 8))
        self.combo_book.pack(side="left", padx=4)
        self.combo_book.bind("<<ComboboxSelected>>", lambda e: self._on_book_selected())

        # 버튼 행
        r1 = tk.Frame(bf); r1.pack(fill="x", padx=6, pady=(2, 4))
        tk.Button(r1, text="＋ 새 단어장", command=self._add_book,
                  bg="#2563eb", fg="white", font=("맑은 고딕", 8, "bold"),
                  relief="flat", padx=8).pack(side="left", padx=2)
        tk.Button(r1, text="✏ 수정", command=self._edit_book,
                  font=("맑은 고딕", 8), relief="flat", padx=6).pack(side="left", padx=2)
        tk.Button(r1, text="🗑 삭제", command=self._delete_book,
                  bg="#dc2626", fg="white", font=("맑은 고딕", 8),
                  relief="flat", padx=6).pack(side="left", padx=2)
        tk.Button(r1, text="📋 AI 추출 가져오기", command=self._open_ai_import,
                  bg="#7c3aed", fg="white", font=("맑은 고딕", 8, "bold"),
                  relief="flat", padx=8).pack(side="right", padx=2)
        tk.Button(r1, text="📁 파일", command=self._import_file,
                  font=("맑은 고딕", 8), relief="flat", padx=6).pack(side="right", padx=2)

        # ── 단원(Day) + 단어 목록 (상하 분할) ──
        content_pw = tk.PanedWindow(parent, orient="vertical", sashwidth=4, bg="#e2e8f0")
        content_pw.pack(fill="both", expand=True, padx=4, pady=2)

        # ─── 단원 (Day) 목록 ───
        uf = tk.LabelFrame(content_pw, text=" 📋 단원 (Day) ", font=("맑은 고딕", 9, "bold"))
        content_pw.add(uf, minsize=100)

        unit_inner = tk.Frame(uf)
        unit_inner.pack(fill="both", expand=True, padx=6, pady=4)

        self.tree_units = ttk.Treeview(unit_inner, columns=("이름", "단어수"),
                                        show="headings", height=4, selectmode="extended")
        self.tree_units.heading("이름", text="단원명")
        self.tree_units.heading("단어수", text="단어수")
        self.tree_units.column("이름", width=180)
        self.tree_units.column("단어수", width=60, anchor="center")
        scr_u = ttk.Scrollbar(unit_inner, orient="vertical", command=self.tree_units.yview)
        self.tree_units.configure(yscrollcommand=scr_u.set)
        self.tree_units.pack(side="left", fill="both", expand=True)
        scr_u.pack(side="right", fill="y")
        self.tree_units.bind("<<TreeviewSelect>>", lambda e: self._on_unit_selected())

        ur = tk.Frame(uf); ur.pack(fill="x", padx=6, pady=(0, 4))
        tk.Button(ur, text="＋ 단원", command=self._add_unit,
                  font=("맑은 고딕", 8), relief="flat", padx=6).pack(side="left", padx=2)
        tk.Button(ur, text="✏", command=self._edit_unit,
                  font=("맑은 고딕", 8), relief="flat", width=3).pack(side="left", padx=1)
        tk.Button(ur, text="🗑", command=self._delete_unit,
                  bg="#dc2626", fg="white", font=("맑은 고딕", 8),
                  relief="flat", width=3).pack(side="left", padx=1)
        tk.Button(ur, text="▶ 단원 전체 → 출제범위", command=self._add_units_to_scope,
                  bg="#16a34a", fg="white", font=("맑은 고딕", 8, "bold"),
                  relief="flat", padx=8).pack(side="right", padx=2)

        # ─── 단어 목록 ───
        wf = tk.LabelFrame(content_pw, text=" 📝 단어 목록 ", font=("맑은 고딕", 9, "bold"))
        content_pw.add(wf, minsize=150)

        # v7.1: selectmode="extended" + 예문유무 컬럼
        word_inner = tk.Frame(wf)
        word_inner.pack(fill="both", expand=True, padx=6, pady=4)

        self.tree_words = ttk.Treeview(word_inner,
                                        columns=("No", "영어", "한국어", "품사", "예문"),
                                        show="headings", height=8, selectmode="extended")
        self.tree_words.heading("No", text="#")
        self.tree_words.heading("영어", text="English")
        self.tree_words.heading("한국어", text="한국어")
        self.tree_words.heading("품사", text="품사")
        self.tree_words.heading("예문", text="예문")
        self.tree_words.column("No", width=30, anchor="center")
        self.tree_words.column("영어", width=120)
        self.tree_words.column("한국어", width=120)
        self.tree_words.column("품사", width=40, anchor="center")
        self.tree_words.column("예문", width=35, anchor="center")
        scr_w = ttk.Scrollbar(word_inner, orient="vertical", command=self.tree_words.yview)
        self.tree_words.configure(yscrollcommand=scr_w.set)
        self.tree_words.pack(side="left", fill="both", expand=True)
        scr_w.pack(side="right", fill="y")

        # v7.1: Ctrl+A 전체선택 바인딩
        self.tree_words.bind("<Control-a>", self._select_all_words)
        self.tree_words.bind("<Control-A>", self._select_all_words)
        self.tree_words.bind("<<TreeviewSelect>>", lambda e: self._on_word_selected())

        # 단어 목록 하단 버튼
        wr = tk.Frame(wf); wr.pack(fill="x", padx=6, pady=(0, 4))
        self.lbl_word_count = tk.Label(wr, text="0개", font=("맑은 고딕", 8), fg="#475569")
        self.lbl_word_count.pack(side="left")

        # v7.1: 선택 단어 → 출제범위 버튼 (핵심 추가)
        tk.Button(wr, text="▶ 선택 단어 → 출제범위", command=self._add_words_to_scope,
                  bg="#0891b2", fg="white", font=("맑은 고딕", 8, "bold"),
                  relief="flat", padx=8).pack(side="right", padx=2)
        tk.Button(wr, text="＋", command=self._add_word,
                  font=("맑은 고딕", 8), relief="flat", width=3).pack(side="right", padx=1)
        tk.Button(wr, text="✏", command=self._edit_word,
                  font=("맑은 고딕", 8), relief="flat", width=3).pack(side="right", padx=1)
        tk.Button(wr, text="🗑", command=self._delete_word,
                  bg="#dc2626", fg="white", font=("맑은 고딕", 8),
                  relief="flat", width=3).pack(side="right", padx=1)

        # ─── v7.0 예문(Sentence) Detail 패널 ───
        sf = tk.LabelFrame(parent, text=" 💬 예문 (선택된 단어) ", font=("맑은 고딕", 9, "bold"))
        sf.pack(fill="x", padx=4, pady=(2, 4))

        sent_inner = tk.Frame(sf)
        sent_inner.pack(fill="both", expand=True, padx=6, pady=4)

        self.tree_sentences = ttk.Treeview(sent_inner,
                                            columns=("난이도", "영어 예문", "한국어 해석", "target"),
                                            show="headings", height=3)
        self.tree_sentences.heading("난이도", text="난이도")
        self.tree_sentences.heading("영어 예문", text="English Sentence")
        self.tree_sentences.heading("한국어 해석", text="한국어 해석")
        self.tree_sentences.heading("target", text="target")
        self.tree_sentences.column("난이도", width=45, anchor="center")
        self.tree_sentences.column("영어 예문", width=200)
        self.tree_sentences.column("한국어 해석", width=140)
        self.tree_sentences.column("target", width=55, anchor="center")
        scr_s = ttk.Scrollbar(sent_inner, orient="vertical", command=self.tree_sentences.yview)
        self.tree_sentences.configure(yscrollcommand=scr_s.set)
        self.tree_sentences.pack(side="left", fill="both", expand=True)
        scr_s.pack(side="right", fill="y")

        sr = tk.Frame(sf); sr.pack(fill="x", padx=6, pady=(0, 4))
        self.lbl_sent_count = tk.Label(sr, text="예문 0개", font=("맑은 고딕", 8), fg="#475569")
        self.lbl_sent_count.pack(side="left")
        tk.Button(sr, text="＋ 예문", command=self._add_sentence_dialog,
                  bg="#7c3aed", fg="white", font=("맑은 고딕", 8, "bold"),
                  relief="flat", padx=6).pack(side="right", padx=2)
        tk.Button(sr, text="✏", command=self._edit_sentence_dialog,
                  font=("맑은 고딕", 8), relief="flat", width=3).pack(side="right", padx=1)
        tk.Button(sr, text="🗑", command=self._delete_sentence,
                  bg="#dc2626", fg="white", font=("맑은 고딕", 8),
                  relief="flat", width=3).pack(side="right", padx=1)

    # ═══════════════════════════════════════════
    # 우측: 출제범위 + 설정 + 생성
    # ═══════════════════════════════════════════

    def _build_right(self, parent):
        # ── 출제 범위 ──
        sf = tk.LabelFrame(parent, text=" 🎯 출제 범위 ", font=("맑은 고딕", 9, "bold"))
        sf.pack(fill="both", expand=True, padx=4, pady=(4, 2))

        self.lbl_scope_info = tk.Label(sf, text="단원 0개 · 단어 0개",
                                        font=("맑은 고딕", 9, "bold"), fg="#2563eb")
        self.lbl_scope_info.pack(padx=6, pady=(4, 2))

        self.list_scope = tk.Listbox(sf, height=8, selectmode=tk.EXTENDED,
                                      font=("맑은 고딕", 9), bg="#f8fafc")
        self.list_scope.pack(fill="both", expand=True, padx=6, pady=2)

        btn_scope = tk.Frame(sf); btn_scope.pack(fill="x", padx=6, pady=(2, 4))
        tk.Button(btn_scope, text="선택 제거", command=self._remove_from_scope,
                  font=("맑은 고딕", 8), relief="flat", padx=6).pack(side="left", padx=2)
        tk.Button(btn_scope, text="전체 비우기", command=self._clear_scope,
                  font=("맑은 고딕", 8), relief="flat", padx=6).pack(side="left", padx=2)

        # ── 시험지 설정 ──
        gf = tk.LabelFrame(parent, text=" 📄 시험지 생성 ", font=("맑은 고딕", 9, "bold"))
        gf.pack(fill="x", padx=4, pady=2)

        # 학생 + 모드 (한 줄)
        r1 = tk.Frame(gf); r1.pack(fill="x", padx=8, pady=3)
        tk.Label(r1, text="학생:", font=("맑은 고딕", 8)).pack(side="left")
        self.combo_vstudent = ttk.Combobox(r1, values=[], width=13, state="readonly",
                                            font=("맑은 고딕", 8))
        self.combo_vstudent.pack(side="left", padx=4)
        tk.Label(r1, text="모드:", font=("맑은 고딕", 8)).pack(side="left", padx=(8, 0))
        self.combo_mode = ttk.Combobox(r1, values=["영어→한국어", "한국어→영어"],
                                        width=11, state="readonly", font=("맑은 고딕", 8))
        self.combo_mode.current(0); self.combo_mode.pack(side="left", padx=4)

        # 문항수 + 옵션 (한 줄)
        r2 = tk.Frame(gf); r2.pack(fill="x", padx=8, pady=3)
        tk.Label(r2, text="문항수:", font=("맑은 고딕", 8)).pack(side="left")
        self.ent_vcount = tk.Entry(r2, width=5, font=("맑은 고딕", 8))
        self.ent_vcount.insert(0, "40")
        self.ent_vcount.pack(side="left", padx=4)
        tk.Label(r2, text="(0=전체)", fg="#94a3b8", font=("맑은 고딕", 7)).pack(side="left")
        self.var_shuffle = tk.IntVar(value=1)
        tk.Checkbutton(r2, text="셔플", variable=self.var_shuffle,
                       font=("맑은 고딕", 8)).pack(side="left", padx=(12, 0))
        self.var_hint = tk.IntVar(value=0)
        tk.Checkbutton(r2, text="첫글자힌트", variable=self.var_hint,
                       font=("맑은 고딕", 8)).pack(side="left", padx=4)

        # 포맷
        r3 = tk.Frame(gf); r3.pack(fill="x", padx=8, pady=3)
        tk.Label(r3, text="포맷:", font=("맑은 고딕", 8)).pack(side="left")
        self.combo_vformat = ttk.Combobox(r3, values=["Word + PDF", "Word만", "PDF만"],
                                           width=10, state="readonly", font=("맑은 고딕", 8))
        self.combo_vformat.current(0); self.combo_vformat.pack(side="left", padx=4)

        # 고급 설정 (접이식)
        self._build_advanced_settings(gf)

        # ── 📝 단어 시험지 생성 버튼 ──
        tk.Button(gf, text="📝 단어 시험지 생성", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 11, "bold"), height=2, relief="flat",
                  command=self._generate_vocab_test).pack(fill="x", padx=8, pady=(6, 4))

        # ── 💬 예문 시험지 ──
        ef = tk.LabelFrame(parent, text=" 💬 예문 시험지 ", font=("맑은 고딕", 9, "bold"),
                           fg="#7c3aed")
        ef.pack(fill="x", padx=4, pady=(2, 4))

        er1 = tk.Frame(ef); er1.pack(fill="x", padx=8, pady=3)
        tk.Label(er1, text="모드:", font=("맑은 고딕", 8)).pack(side="left")
        self.combo_sent_mode = ttk.Combobox(er1,
                                             values=["문장해석 (영→한)", "빈칸영작 (Cloze)"],
                                             width=16, state="readonly", font=("맑은 고딕", 8))
        self.combo_sent_mode.current(0); self.combo_sent_mode.pack(side="left", padx=4)
        tk.Label(er1, text="난이도:", font=("맑은 고딕", 8)).pack(side="left", padx=(8, 0))
        self.combo_sent_diff = ttk.Combobox(er1, values=["전체", "1-기본", "2-핵심", "3-심화"],
                                             width=7, state="readonly", font=("맑은 고딕", 8))
        self.combo_sent_diff.current(0); self.combo_sent_diff.pack(side="left", padx=4)

        er2 = tk.Frame(ef); er2.pack(fill="x", padx=8, pady=1)
        self.var_sent_hint = tk.IntVar(value=0)
        tk.Checkbutton(er2, text="첫글자 힌트", variable=self.var_sent_hint,
                       font=("맑은 고딕", 8)).pack(side="left")

        tk.Button(ef, text="📝 예문 시험지 생성", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"), relief="flat",
                  command=self._generate_sentence_test).pack(fill="x", padx=8, pady=(4, 6))

    def _build_advanced_settings(self, parent):
        """접이식 고급 설정"""
        self._adv_visible = tk.BooleanVar(value=False)
        toggle_f = tk.Frame(parent)
        toggle_f.pack(fill="x", padx=8, pady=1)
        self._adv_toggle_btn = tk.Button(toggle_f, text="▸ 고급 설정",
                                          font=("맑은 고딕", 8), fg="#6b7280",
                                          relief="flat", command=self._toggle_advanced)
        self._adv_toggle_btn.pack(side="left")

        self._adv_frame = tk.Frame(parent)
        # 제목
        tf = tk.Frame(self._adv_frame); tf.pack(fill="x", padx=8, pady=1)
        tk.Label(tf, text="제목:", font=("맑은 고딕", 8)).pack(side="left")
        self.ent_vtitle = tk.Entry(tf, width=20, font=("맑은 고딕", 8))
        self.ent_vtitle.pack(side="left", padx=4)
        # 파일명
        ff = tk.Frame(self._adv_frame); ff.pack(fill="x", padx=8, pady=1)
        tk.Label(ff, text="파일명:", font=("맑은 고딕", 8)).pack(side="left")
        self.ent_vfilename = tk.Entry(ff, width=20, font=("맑은 고딕", 8))
        self.ent_vfilename.pack(side="left", padx=4)
        # 배너
        bnf = tk.Frame(self._adv_frame); bnf.pack(fill="x", padx=8, pady=1)
        tk.Button(bnf, text="🖼 배너", font=("맑은 고딕", 8), relief="flat",
                  command=self._select_banner).pack(side="left")
        tk.Button(bnf, text="✕", font=("맑은 고딕", 8), width=2, relief="flat",
                  command=self._remove_banner).pack(side="left", padx=2)
        self.lbl_vbanner = tk.Label(bnf, text="(없음)", fg="#94a3b8", font=("맑은 고딕", 7))
        self.lbl_vbanner.pack(side="left", padx=2)

        tk.Label(self._adv_frame, text="비워두면 기본값 자동 적용",
                 font=("맑은 고딕", 7), fg="#94a3b8").pack(padx=8, pady=1)

    def _toggle_advanced(self):
        if self._adv_visible.get():
            self._adv_frame.pack_forget()
            self._adv_toggle_btn.config(text="▸ 고급 설정")
            self._adv_visible.set(False)
        else:
            self._adv_frame.pack(fill="x", after=self._adv_toggle_btn.master)
            self._adv_toggle_btn.config(text="▾ 고급 설정")
            self._adv_visible.set(True)

    # ═══════════════════════════════════════════
    # v7.1: Ctrl+A 전체선택
    # ═══════════════════════════════════════════

    def _select_all_words(self, event=None):
        """Ctrl+A: 단어 목록 전체 선택"""
        children = self.tree_words.get_children()
        if children:
            self.tree_words.selection_set(children)
        return "break"  # 기본 동작 방지

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

    def _get_all_categories(self):
        """DB에서 실제 사용중인 카테고리 + 기본 카테고리 합치기"""
        books = list_vocab_books(self.cfg)
        db_cats = sorted(set(b.get("category", "기본단어장") for b in books if b.get("category")))
        all_cats = list(VOCAB_CATEGORIES)
        for c in db_cats:
            if c not in all_cats:
                all_cats.append(c)
        return all_cats

    def _on_category_changed(self):
        """v7.1: 카테고리 필터 변경 시 단어장 목록 갱신"""
        self._refresh_book_list()

    def _refresh_book_list(self):
        """카테고리 필터 적용하여 단어장 콤보 갱신"""
        cat = self.combo_category.get()
        books = list_vocab_books(self.cfg)
        if cat and cat != "전체":
            books = [b for b in books if b.get("category", "") == cat]
        self.combo_book["values"] = [f"[{b['id']}] {b['name']} ({b.get('category','')})" for b in books]
        if books:
            self.combo_book.current(0)
            self._on_book_selected()
        else:
            self.combo_book.set("")
            self.tree_units.delete(*self.tree_units.get_children())
            self.tree_words.delete(*self.tree_words.get_children())

    def _add_book(self):
        """v7.1: 개선된 새 단어장 다이얼로그"""
        all_cats = self._get_all_categories()

        top = tk.Toplevel(self.app.root)
        top.title("📖 새 단어장 만들기")
        top.geometry("420x280")
        top.attributes("-topmost", True)
        top.configure(bg="#f8fafc")

        # 헤더
        hdr = tk.Frame(top, bg="#2563eb", height=40)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="📖 새 단어장 만들기", font=("맑은 고딕", 11, "bold"),
                 bg="#2563eb", fg="white").pack(pady=8)

        body = tk.Frame(top, bg="#f8fafc")
        body.pack(fill="both", expand=True, padx=20, pady=12)

        # 카테고리
        tk.Label(body, text="카테고리", font=("맑은 고딕", 9, "bold"),
                 bg="#f8fafc", fg="#334155").pack(anchor="w", pady=(0, 2))
        cat_f = tk.Frame(body, bg="#f8fafc"); cat_f.pack(fill="x", pady=(0, 10))
        combo_cat = ttk.Combobox(cat_f, values=all_cats, width=18, font=("맑은 고딕", 9))
        combo_cat.set("기본단어장")
        combo_cat.pack(side="left")
        tk.Label(cat_f, text="← 직접 입력 가능 (새 카테고리 자동 기억)",
                 fg="#94a3b8", font=("맑은 고딕", 8), bg="#f8fafc").pack(side="left", padx=6)

        # 단어장 이름
        tk.Label(body, text="단어장 이름", font=("맑은 고딕", 9, "bold"),
                 bg="#f8fafc", fg="#334155").pack(anchor="w", pady=(0, 2))
        ent_name = tk.Entry(body, width=35, font=("맑은 고딕", 10))
        ent_name.pack(fill="x", pady=(0, 10), ipady=3)
        ent_name.focus_set()

        # 설명
        tk.Label(body, text="설명 (선택)", font=("맑은 고딕", 8),
                 bg="#f8fafc", fg="#94a3b8").pack(anchor="w", pady=(0, 2))
        ent_desc = tk.Entry(body, width=35, font=("맑은 고딕", 9), fg="#64748b")
        ent_desc.pack(fill="x", pady=(0, 8))

        def do_add():
            name = ent_name.get().strip()
            if not name:
                return messagebox.showwarning("오류", "단어장 이름을 입력하세요.", parent=top)
            cat = combo_cat.get().strip() or "기본단어장"
            add_vocab_book(self.cfg, name, cat)
            top.destroy()
            # 새로 만든 카테고리 반영
            self._refresh_category_combo()
            self.combo_category.set(cat)
            self.refresh()

        btn_f = tk.Frame(body, bg="#f8fafc"); btn_f.pack(fill="x", pady=(4, 0))
        tk.Button(btn_f, text="취소", command=top.destroy,
                  font=("맑은 고딕", 9), relief="flat", padx=12).pack(side="left")
        tk.Button(btn_f, text="✔ 만들기", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 10, "bold"), relief="flat", padx=16,
                  command=do_add).pack(side="right")
        ent_name.bind("<Return>", lambda e: do_add())

    def _refresh_category_combo(self):
        """카테고리 콤보 갱신 (DB에서 사용중인 카테고리 포함)"""
        all_cats = self._get_all_categories()
        current = self.combo_category.get()
        self.combo_category["values"] = ["전체"] + all_cats
        if current in (["전체"] + all_cats):
            self.combo_category.set(current)
        else:
            self.combo_category.current(0)

    def _edit_book(self):
        bid = self._get_current_book_id()
        if not bid:
            return messagebox.showwarning("알림", "단어장을 선택하세요.")
        book = get_vocab_book(self.cfg, bid)
        if not book:
            return

        all_cats = self._get_all_categories()
        top = tk.Toplevel(self.app.root); top.title("단어장 수정"); top.geometry("400x220"); top.attributes("-topmost", True)
        top.configure(bg="#f8fafc")

        body = tk.Frame(top, bg="#f8fafc"); body.pack(fill="both", expand=True, padx=20, pady=16)

        tk.Label(body, text="이름:", font=("맑은 고딕", 9, "bold"), bg="#f8fafc").pack(anchor="w")
        ent_name = tk.Entry(body, width=30, font=("맑은 고딕", 10))
        ent_name.pack(fill="x", pady=(2, 10), ipady=2)
        ent_name.insert(0, book["name"])

        tk.Label(body, text="카테고리:", font=("맑은 고딕", 9, "bold"), bg="#f8fafc").pack(anchor="w")
        combo_cat = ttk.Combobox(body, values=all_cats, width=18, font=("맑은 고딕", 9))
        combo_cat.set(book.get("category", "기본단어장"))
        combo_cat.pack(anchor="w", pady=(2, 12))

        def do_save():
            update_vocab_book(self.cfg, bid, name=ent_name.get().strip(),
                              category=combo_cat.get().strip())
            top.destroy()
            self._refresh_category_combo()
            self.refresh()

        tk.Button(body, text="저장", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 10, "bold"), relief="flat",
                  command=do_save).pack(fill="x", pady=4)

    def _delete_book(self):
        bid = self._get_current_book_id()
        if not bid:
            return messagebox.showwarning("알림", "단어장을 선택하세요.")
        if not messagebox.askyesno("경고", "이 단어장을 삭제하시겠습니까?\n(학습 이력은 보존됩니다)"):
            return
        delete_vocab_book(self.cfg, bid)
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
        top = tk.Toplevel(self.app.root); top.title("단어 추가"); top.geometry("400x220"); top.attributes("-topmost", True)
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
                return messagebox.showwarning("오류", "영어 단어를 입력하세요.", parent=top)
            add_vocab_word(self.cfg, uid, eng, kor, pos)
            top.destroy()
            self._refresh_words()

        tk.Button(top, text="추가", bg="#2563eb", fg="white", font=("맑은 고딕", 10, "bold"),
                  relief="flat", command=do_add).pack(fill="x", padx=16, pady=12)
        entries["영어:"].bind("<Return>", lambda e: do_add())

    def _edit_word(self):
        sel = self.tree_words.selection()
        if not sel:
            return messagebox.showwarning("알림", "단어를 선택하세요.")
        wid = int(sel[0])
        vals = self.tree_words.item(sel[0], "values")

        top = tk.Toplevel(self.app.root); top.title("단어 수정"); top.geometry("400x220"); top.attributes("-topmost", True)
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

        tk.Button(top, text="저장", bg="#16a34a", fg="white", font=("맑은 고딕", 10, "bold"),
                  relief="flat", command=do_save).pack(fill="x", padx=16, pady=12)

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
    # v7.1: 출제 범위 관리 (단원 + 개별 단어)
    # ═══════════════════════════════════════════

    def _add_units_to_scope(self):
        """선택된 단원 전체를 출제 범위에 추가"""
        sel = self.tree_units.selection()
        if not sel:
            return messagebox.showwarning("알림", "단원을 선택하세요.")
        added = 0
        for s in sel:
            uid = int(s)
            if uid not in self.selected_units:
                self.selected_units.append(uid)
                added += 1
        if added:
            self._refresh_scope()

    def _add_words_to_scope(self):
        """v7.1: 선택된 개별 단어를 출제 범위에 추가"""
        sel = self.tree_words.selection()
        if not sel:
            return messagebox.showwarning("알림", "단어를 선택하세요.\n(Ctrl+A로 전체선택, Ctrl+클릭으로 추가선택)")
        added = 0
        for s in sel:
            wid = int(s)
            if wid not in self.selected_word_ids:
                self.selected_word_ids.append(wid)
                added += 1
        if added:
            self._refresh_scope()

    def _remove_from_scope(self):
        """출제범위 리스트에서 선택 항목 제거"""
        indices = list(self.list_scope.curselection())
        if not indices:
            return
        # 역순으로 삭제 (인덱스 깨짐 방지)
        unit_count = len(self.selected_units)
        for idx in reversed(indices):
            if idx < unit_count:
                self.selected_units.pop(idx)
            else:
                word_idx = idx - unit_count
                if word_idx < len(self.selected_word_ids):
                    self.selected_word_ids.pop(word_idx)
        self._refresh_scope()

    def _clear_scope(self):
        self.selected_units.clear()
        self.selected_word_ids.clear()
        self._refresh_scope()

    def _refresh_scope(self):
        """출제범위 리스트박스 갱신 (단원 + 개별 단어)"""
        self.list_scope.delete(0, tk.END)
        total_words = 0

        # 1) 단원별 항목
        for uid in self.selected_units:
            from exam_bank.models.vocabulary import get_vocab_unit
            unit = get_vocab_unit(self.cfg, uid)
            if unit:
                book = get_vocab_book(self.cfg, unit.get("book_id"))
                book_name = book["name"] if book else "?"
                wc = unit.get("word_count", 0)
                total_words += wc
                self.list_scope.insert("end", f"📋 {book_name} > {unit['unit_name']} ({wc}단어)")

        # 2) 개별 단어 항목
        if self.selected_word_ids:
            self.list_scope.insert("end", f"── 개별 선택 단어 ({len(self.selected_word_ids)}개) ──")
            total_words += len(self.selected_word_ids)

        info_parts = []
        if self.selected_units:
            info_parts.append(f"단원 {len(self.selected_units)}개")
        if self.selected_word_ids:
            info_parts.append(f"개별단어 {len(self.selected_word_ids)}개")
        info_parts.append(f"총 {total_words}단어")
        self.lbl_scope_info.config(text=" · ".join(info_parts) if info_parts else "단원 0개 · 단어 0개")

    # ═══════════════════════════════════════════
    # AI 추출 가져오기
    # ═══════════════════════════════════════════

    def _open_ai_import(self):
        bid = self._get_current_book_id()
        if not bid:
            return messagebox.showwarning("알림", "먼저 단어장을 선택하세요.")

        top = tk.Toplevel(self.app.root)
        top.title("AI 추출 가져오기 (v7.1)")
        top.geometry("750x720")
        top.attributes("-topmost", True)

        # ── 1️⃣ 프롬프트 선택 + 복사 ──
        pf = tk.LabelFrame(top, text=" 1️⃣ AI 프롬프트 (사진과 함께 붙여넣기) ", font=("맑은 고딕", 9, "bold"))
        pf.pack(fill="x", padx=12, pady=6)

        mode_f = tk.Frame(pf); mode_f.pack(fill="x", padx=8, pady=2)
        prompt_mode = tk.IntVar(value=1)
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
            messagebox.showinfo("복사 완료", "프롬프트가 클립보드에 복사되었습니다.\nClaude/ChatGPT 채팅에 사진과 함께 붙여넣으세요.", parent=top)

        tk.Button(pf, text="📋 프롬프트 복사", command=copy_prompt,
                  bg="#2563eb", fg="white", font=("맑은 고딕", 9, "bold"), relief="flat").pack(padx=8, pady=2)

        # ── 2️⃣ JSON 붙여넣기 ──
        jf = tk.LabelFrame(top, text=" 2️⃣ AI가 출력한 JSON 붙여넣기 ", font=("맑은 고딕", 9, "bold"))
        jf.pack(fill="x", padx=12, pady=6)
        txt_json = tk.Text(jf, height=7, font=("Consolas", 9))
        txt_json.pack(fill="x", padx=8, pady=4)

        df = tk.Frame(jf); df.pack(fill="x", padx=8, pady=4)
        tk.Label(df, text="Day 이름:").pack(side="left")
        ent_day = tk.Entry(df, width=15); ent_day.pack(side="left", padx=4)
        tk.Label(df, text="(비우면 section 필드로 자동 분류)", fg="#94a3b8", font=("맑은 고딕", 8)).pack(side="left")

        # ── 3️⃣ 미리보기 ──
        pvf = tk.LabelFrame(top, text=" 3️⃣ 미리보기 ", font=("맑은 고딕", 9, "bold"))
        pvf.pack(fill="both", expand=True, padx=12, pady=6)

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
                return messagebox.showwarning("오류", "JSON을 붙여넣으세요.", parent=top)

            raw = re.sub(r'^```(?:json)?\s*', '', raw).strip()
            raw = re.sub(r'```\s*$', '', raw).strip()

            try:
                data = json.loads(raw)
            except json.JSONDecodeError as e:
                return messagebox.showerror("JSON 오류", f"JSON 파싱 실패:\n{e}\n\n앞뒤 마크다운 블록(```)이 포함되었는지 확인하세요.", parent=top)

            if not isinstance(data, list):
                return messagebox.showerror("형식 오류", "JSON 배열 형식이어야 합니다.\n[{...}, {...}, ...]", parent=top)

            parsed_words = data
            tree_pv.delete(*tree_pv.get_children())
            total_sents = 0
            for i, w in enumerate(data, 1):
                eng = w.get("english", w.get("word", ""))
                kor = w.get("korean", w.get("meaning", w.get("korean_meaning", "")))
                pos = w.get("pos", w.get("part_of_speech", ""))
                sents = w.get("sentences", [])
                sc = len(sents)
                has_example = bool(w.get("example", w.get("example_sentence", "")).strip())
                if has_example:
                    sc += 1
                total_sents += sc
                tree_pv.insert("", "end", values=(i, eng, kor, pos, sc if sc else "-"))

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
                return messagebox.showwarning("오류", "먼저 '미리보기'를 클릭하세요.", parent=top)
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

            sent_total = results.pop("__sentences_total__", 0)
            total = sum(results.values())
            detail = "\n".join(f"  {k}: {v}개" for k, v in results.items())
            sent_msg = f"\n📝 예문 {sent_total}개 자동저장" if sent_total else ""
            messagebox.showinfo("가져오기 완료", f"총 {total}개 단어 등록!{sent_msg}\n\n{detail}", parent=top)
            top.destroy()
            self.refresh()

        btn_f = tk.Frame(top); btn_f.pack(fill="x", padx=12, pady=8)
        tk.Button(btn_f, text="👁 미리보기", command=do_parse,
                  bg="#6b7280", fg="white", font=("맑은 고딕", 9, "bold"), relief="flat").pack(side="left", padx=4)
        tk.Button(btn_f, text="✅ DB에 저장", command=do_import,
                  bg="#16a34a", fg="white", font=("맑은 고딕", 10, "bold"), relief="flat").pack(side="right", padx=4)

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

    def _collect_all_test_words(self):
        """출제범위(단원+개별단어)에서 단어 풀 수집"""
        all_words = []
        seen_ids = set()

        # 1) 단원별 단어
        if self.selected_units:
            unit_words = get_words_by_unit_ids(self.cfg, self.selected_units)
            for w in unit_words:
                if w["id"] not in seen_ids:
                    all_words.append(w)
                    seen_ids.add(w["id"])

        # 2) 개별 선택 단어
        if self.selected_word_ids:
            for wid in self.selected_word_ids:
                if wid not in seen_ids:
                    # DB에서 개별 조회
                    from exam_bank.models.vocabulary import get_vocab_unit
                    from exam_bank.models.database import db_conn
                    with db_conn(self.cfg) as conn:
                        cur = conn.cursor()
                        row = cur.execute(
                            """SELECT w.*, u.unit_name, u.book_id, b.name as book_name
                               FROM vocab_words w
                               JOIN vocab_units u ON w.unit_id = u.id
                               JOIN vocab_books b ON u.book_id = b.id
                               WHERE w.id=?""", (wid,)).fetchone()
                        if row:
                            cols = [d[0] for d in cur.description]
                            all_words.append(dict(zip(cols, row)))
                            seen_ids.add(wid)

        return all_words

    def _generate_vocab_test(self):
        if not self.selected_units and not self.selected_word_ids:
            return messagebox.showwarning("경고", "출제 범위를 먼저 설정하세요.\n\n• 단원 선택 → '▶ 단원 전체 → 출제범위'\n• 단어 선택 → '▶ 선택 단어 → 출제범위'")

        all_words = self._collect_all_test_words()
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

        exam_type = "eng_to_kor" if self.combo_mode.current() == 0 else "kor_to_eng"

        sid = self._get_vstudent_id()
        student_name = ""
        if sid:
            s = get_student(self.cfg, sid)
            student_name = s["name"] if s else ""

        unit_names = []
        for uid in self.selected_units:
            from exam_bank.models.vocabulary import get_vocab_unit
            u = get_vocab_unit(self.cfg, uid)
            if u:
                unit_names.append(u["unit_name"])
        if self.selected_word_ids:
            unit_names.append(f"개별선택{len(self.selected_word_ids)}단어")

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
                custom_title=self.ent_vtitle.get().strip() if self._adv_visible.get() else "",
                custom_filename=self.ent_vfilename.get().strip() if self._adv_visible.get() else "",
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

    def _generate_sentence_test(self):
        """v7.1: 예문 시험지 생성 (버그 수정: dict→list 변환)"""
        if not self.selected_units and not self.selected_word_ids:
            return messagebox.showwarning("경고", "출제 범위를 먼저 설정하세요.")

        all_words = self._collect_all_test_words()
        if not all_words:
            return messagebox.showwarning("경고", "선택된 범위에 단어가 없습니다.")

        word_ids = [w["id"] for w in all_words]
        word_map = {w["id"]: w for w in all_words}

        # ★ v7.1 버그수정: get_sentences_by_word_ids는 dict 반환
        # {word_id: [sentence_dict, ...]} → 리스트로 평탄화
        sentences_dict = get_sentences_by_word_ids(self.cfg, word_ids)
        if not sentences_dict:
            return messagebox.showwarning("경고", "선택된 범위에 예문이 없습니다.\nAI 가져오기에서 '단어 + 예문' 모드로 먼저 추가하세요.")

        # dict → flat list 변환
        all_sentences = []
        for wid, sent_list in sentences_dict.items():
            for s in sent_list:
                all_sentences.append(s)

        # 난이도 필터
        diff_sel = self.combo_sent_diff.get()
        if diff_sel.startswith("1"):
            all_sentences = [s for s in all_sentences if s.get("difficulty", 1) == 1]
        elif diff_sel.startswith("2"):
            all_sentences = [s for s in all_sentences if s.get("difficulty", 1) == 2]
        elif diff_sel.startswith("3"):
            all_sentences = [s for s in all_sentences if s.get("difficulty", 1) == 3]

        if not all_sentences:
            return messagebox.showwarning("경고", f"난이도 '{diff_sel}' 예문이 없습니다.")

        # 예문 아이템 조합 (word 정보 포함)
        items = []
        for s in all_sentences:
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

        test_mode = "translation" if self.combo_sent_mode.current() == 0 else "cloze"

        sid = self._get_vstudent_id()
        student_name = ""
        if sid:
            s = get_student(self.cfg, sid)
            student_name = s["name"] if s else ""

        unit_names = []
        for uid in self.selected_units:
            from exam_bank.models.vocabulary import get_vocab_unit
            u = get_vocab_unit(self.cfg, uid)
            if u:
                unit_names.append(u["unit_name"])

        mode_label = "문장해석" if test_mode == "translation" else "빈칸영작"

        try:
            files, err = create_sentence_test(
                target_dir=self.cfg["last_dir"],
                sentence_items=items,
                cfg=self.cfg,
                test_mode=test_mode,
                student_name=student_name,
                show_hint=bool(self.var_sent_hint.get()),
                shuffle=bool(self.var_shuffle.get()),
                custom_title=self.ent_vtitle.get().strip() if self._adv_visible.get() else "",
                custom_filename=self.ent_vfilename.get().strip() if self._adv_visible.get() else "",
                logo_path=self.banner_path or self.cfg.get("last_logo", ""),
                output_format=self.combo_vformat.get(),
                unit_names=unit_names,
            )
            if err:
                messagebox.showwarning("경고", f"생성 완료 (일부 오류):\n{err}")
            else:
                fnames = "\n".join(os.path.basename(f) for f in files)
                messagebox.showinfo("완료", f"예문 시험지 생성 완료!\n{len(items)}문항 ({mode_label})\n\n{fnames}\n\n저장: {self.cfg['last_dir']}")
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
        # 카테고리 콤보 갱신
        self._refresh_category_combo()

        # 단어장 목록
        self._refresh_book_list()

        # 학생 목록
        students = list_students(self.cfg)
        self.combo_vstudent["values"] = ["(선택안함)"] + [f"[{s['id']}] {s['name']}" for s in students]
        if self.combo_vstudent["values"] and not self.combo_vstudent.get():
            self.combo_vstudent.current(0)

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
        self.tree_sentences.delete(*self.tree_sentences.get_children())
        self.lbl_sent_count.config(text="예문 0개")
        self._word_sentence_cache.clear()

        uid = self._get_current_unit_id()
        if not uid:
            self.lbl_word_count.config(text="0개")
            return
        words = list_vocab_words(self.cfg, uid)

        # v7.1: 예문 유무 일괄 조회
        word_ids = [w["id"] for w in words]
        if word_ids:
            sent_map = get_sentences_by_word_ids(self.cfg, word_ids)
        else:
            sent_map = {}

        for i, w in enumerate(words, 1):
            wid = w["id"]
            has_sent = "✓" if wid in sent_map and sent_map[wid] else ""
            # example_sentence 필드도 체크
            if not has_sent and w.get("example_sentence", "").strip():
                has_sent = "✓"
            self._word_sentence_cache[wid] = bool(has_sent)
            self.tree_words.insert("", "end", iid=str(wid),
                                   values=(i, w["english"], w["korean"],
                                           w.get("part_of_speech", ""), has_sent))
        self.lbl_word_count.config(text=f"{len(words)}개")

    def _on_word_selected(self):
        """단어 클릭 → 예문 패널 새로고침"""
        self._refresh_sentences()

    def _refresh_sentences(self):
        """선택된 단어의 예문 목록 새로고침"""
        self.tree_sentences.delete(*self.tree_sentences.get_children())
        sel = self.tree_words.selection()
        if not sel:
            self.lbl_sent_count.config(text="예문 0개")
            return
        # 마지막 선택된 단어의 예문 표시
        wid = int(sel[-1])
        sents = list_sentences(self.cfg, wid)
        for s in sents:
            diff_label = SENTENCE_DIFF_SHORT.get(s.get("difficulty", 1), "?")
            en = s.get("sentence_en", "")
            ko = s.get("sentence_ko", "")
            tgt = s.get("target_form", "")
            self.tree_sentences.insert("", "end", iid=str(s["id"]),
                                        values=(diff_label, en, ko, tgt))
        self.lbl_sent_count.config(text=f"예문 {len(sents)}개")

    # ═══════════════════════════════════════════
    # v7.0: 예문 (Sentence) CRUD
    # ═══════════════════════════════════════════

    def _get_selected_word_id(self):
        sel = self.tree_words.selection()
        return int(sel[-1]) if sel else None

    def _get_selected_sentence_id(self):
        sel = self.tree_sentences.selection()
        return int(sel[0]) if sel else None

    def _add_sentence_dialog(self):
        wid = self._get_selected_word_id()
        if not wid:
            return messagebox.showwarning("알림", "먼저 단어를 선택하세요.")

        sel = self.tree_words.selection()
        word_vals = self.tree_words.item(sel[-1], "values") if sel else None
        word_eng = word_vals[1] if word_vals else ""

        top = tk.Toplevel(self.app.root)
        top.title(f"예문 추가 — {word_eng}")
        top.geometry("550x340")
        top.attributes("-topmost", True)

        f1 = tk.Frame(top); f1.pack(fill="x", padx=16, pady=(12, 4))
        tk.Label(f1, text="영어 예문:", width=10, anchor="w").pack(side="left")
        ent_en = tk.Entry(f1, width=45, font=("맑은 고딕", 9)); ent_en.pack(side="left", padx=4)
        ent_en.focus_set()

        f2 = tk.Frame(top); f2.pack(fill="x", padx=16, pady=4)
        tk.Label(f2, text="한국어 해석:", width=10, anchor="w").pack(side="left")
        ent_ko = tk.Entry(f2, width=45, font=("맑은 고딕", 9)); ent_ko.pack(side="left", padx=4)

        f3 = tk.Frame(top); f3.pack(fill="x", padx=16, pady=4)
        tk.Label(f3, text="target form:", width=10, anchor="w").pack(side="left")
        ent_tgt = tk.Entry(f3, width=20, font=("맑은 고딕", 9)); ent_tgt.pack(side="left", padx=4)
        ent_tgt.insert(0, word_eng)
        tk.Label(f3, text="(문장 내 실제 활용형)", fg="#94a3b8", font=("맑은 고딕", 8)).pack(side="left")

        f4 = tk.Frame(top); f4.pack(fill="x", padx=16, pady=4)
        tk.Label(f4, text="난이도:", width=10, anchor="w").pack(side="left")
        diff_vals = [f"{k} - {v}" for k, v in SENTENCE_DIFFICULTY.items()]
        combo_diff = ttk.Combobox(f4, values=diff_vals, width=18, state="readonly")
        combo_diff.current(0); combo_diff.pack(side="left", padx=4)

        f5 = tk.Frame(top); f5.pack(fill="x", padx=16, pady=4)
        tk.Label(f5, text="출처:", width=10, anchor="w").pack(side="left")
        combo_src = ttk.Combobox(f5, values=SENTENCE_SOURCES, width=18)
        combo_src.set("직접입력"); combo_src.pack(side="left", padx=4)

        def do_add():
            en = ent_en.get().strip()
            if not en:
                return messagebox.showwarning("오류", "영어 예문을 입력하세요.", parent=top)
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
            self._refresh_words()  # 예문 유무 컬럼 갱신

        tk.Button(top, text="추가", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"), relief="flat",
                  command=do_add).pack(fill="x", padx=16, pady=12)
        ent_en.bind("<Return>", lambda e: ent_ko.focus_set())
        ent_ko.bind("<Return>", lambda e: do_add())

    def _edit_sentence_dialog(self):
        sid = self._get_selected_sentence_id()
        if not sid:
            return messagebox.showwarning("알림", "예문을 선택하세요.")

        vals = self.tree_sentences.item(str(sid), "values")
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
                  font=("맑은 고딕", 10, "bold"), relief="flat",
                  command=do_save).pack(fill="x", padx=16, pady=12)

    def _delete_sentence(self):
        sid = self._get_selected_sentence_id()
        if not sid:
            return messagebox.showwarning("알림", "예문을 선택하세요.")
        if not messagebox.askyesno("확인", "이 예문을 삭제하시겠습니까?"):
            return
        delete_sentence(self.cfg, sid)
        self._refresh_sentences()
        self._refresh_words()  # 예문 유무 컬럼 갱신

    # ═══════════════════════════════════════════
    # 유틸리티
    # ═══════════════════════════════════════════

    def _simple_input(self, title, prompt):
        top = tk.Toplevel(self.app.root); top.title(title); top.geometry("320x120"); top.attributes("-topmost", True)
        tk.Label(top, text=prompt).pack(pady=(12, 4), padx=16, anchor="w")
        ent = tk.Entry(top, width=30, font=("맑은 고딕", 10)); ent.pack(padx=16, fill="x", ipady=2); ent.focus_set()
        result = [None]

        def do_ok():
            val = ent.get().strip()
            if val:
                result[0] = val
            top.destroy()

        tk.Button(top, text="확인", command=do_ok, bg="#2563eb", fg="white",
                  font=("맑은 고딕", 9, "bold"), relief="flat").pack(fill="x", padx=16, pady=8)
        ent.bind("<Return>", lambda e: do_ok())
        top.wait_window()
        return result[0]
