"""
문제 관리 탭 v6.0
— 드래그 다중 선택 (디바운싱), Ctrl+A (포커스 체크)
— 시험지 제목/파일명 사용자 지정 (sanitization + PermissionError 방어)
— 배너 삽입 UI
— [ ⚙️ 출력 고급 설정 ] LabelFrame 그룹화
"""

import os
import re
import json
import tkinter as tk
from tkinter import messagebox, filedialog, ttk

from exam_bank.constants import SCHOOL_LEVELS, EXAM_TYPES, EXAM_YEARS, DIFFICULTY_LABELS
from exam_bank.models.database import db_conn
from exam_bank.models.passage import (
    search_passages, get_questions_for_passage, delete_passage, delete_question,
    update_passage, get_passage, get_question,
    update_question_answer, update_question_difficulty, auto_select_questions,
)
from exam_bank.models.student import (
    list_students, get_student, get_student_question_stats, create_exam_record,
)
from exam_bank.services.generator import create_exam_files
from exam_bank.services.auto_selector import build_exam_data_from_cart, generate_multi_sets
from exam_bank.config import open_directory
from exam_bank.views.widgets import ToolTip


class BankTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self.cart = {}
        self.preview_data = {}
        self.tooltip = ToolTip(app.root)
        self.banner_path = ""          # v6: 배너 경로
        self._drag_last_row = None     # v6: 드래그 디바운싱
        self._build()

    def _build(self):
        # ── 검색 바 ──
        sf = tk.Frame(self.frame)
        sf.pack(fill="x", padx=10, pady=6)

        tk.Label(sf, text="검색:").pack(side="left")
        self.ent_search = tk.Entry(sf, width=22)
        self.ent_search.pack(side="left", padx=4)
        self.ent_search.bind("<Return>", lambda e: self.refresh())

        tk.Label(sf, text="학교급:").pack(side="left", padx=(10, 0))
        self.f_cat1 = ttk.Combobox(sf, values=["전체"] + SCHOOL_LEVELS, width=6, state="readonly")
        self.f_cat1.current(0); self.f_cat1.pack(side="left", padx=2)

        tk.Label(sf, text="유형:").pack(side="left", padx=(6, 0))
        self.f_cat2 = ttk.Combobox(sf, values=["전체"] + EXAM_TYPES, width=7, state="readonly")
        self.f_cat2.current(0); self.f_cat2.pack(side="left", padx=2)

        tk.Label(sf, text="년도:").pack(side="left", padx=(6, 0))
        self.f_year = ttk.Combobox(sf, values=["전체"] + EXAM_YEARS, width=7, state="readonly")
        self.f_year.current(0); self.f_year.pack(side="left", padx=2)

        tk.Button(sf, text="조회", command=self.refresh, bg="#6b7280", fg="white").pack(side="left", padx=8)
        tk.Button(sf, text="초기화", command=self._reset_filters).pack(side="left")

        # ── 트리뷰 ──
        cols = ("ID", "구분", "분류/시기", "출판사/태그", "내용 요약", "난이도", "등록일")
        self.tree = ttk.Treeview(self.frame, columns=cols, show="headings", height=10, selectmode="extended")
        for c, w in zip(cols, [55, 45, 130, 130, 370, 50, 80]):
            self.tree.heading(c, text=c); self.tree.column(c, width=w)
        self.tree.pack(fill="x", padx=10, pady=4)
        self.tree.bind("<Motion>", self._on_motion)
        self.tree.bind("<Leave>", lambda e: self.tooltip.hide())
        self.tree.bind("<Double-1>", lambda e: self._cart_add())

        # v6: 드래그 선택 (디바운싱)
        self.tree.bind("<B1-Motion>", self._on_drag_select)
        self.tree.bind("<ButtonRelease-1>", self._on_drag_end)

        # v6: Ctrl+A 전체 선택 (트리뷰 포커스일 때만)
        self.tree.bind("<Control-a>", self._select_all_tree)
        self.tree.bind("<Control-A>", self._select_all_tree)

        # ── 트리 액션 ──
        tb = tk.Frame(self.frame); tb.pack(fill="x", padx=10, pady=2)
        tk.Button(tb, text="선택 수정", command=self._edit_selected).pack(side="left", padx=2)
        tk.Button(tb, text="정답/난이도 설정", command=self._set_answer, bg="#0891b2", fg="white").pack(side="left", padx=2)
        tk.Button(tb, text="DB에서 삭제", bg="#dc2626", fg="white", font=("맑은 고딕", 9, "bold"), command=self._delete_selected).pack(side="left", padx=10)
        tk.Button(tb, text="선택 항목 장바구니 담기", bg="#2563eb", fg="white", font=("맑은 고딕", 9, "bold"), command=self._cart_add).pack(side="right", padx=2)

        # ── 자동 출제 ──
        af = tk.LabelFrame(self.frame, text=" 난이도별 자동 출제 ", font=("맑은 고딕", 9, "bold"))
        af.pack(fill="x", padx=10, pady=4)
        ar = tk.Frame(af); ar.pack(fill="x", padx=10, pady=6)
        tk.Label(ar, text="문항수:").pack(side="left")
        self.ent_auto_count = tk.Entry(ar, width=4); self.ent_auto_count.insert(0, "15"); self.ent_auto_count.pack(side="left", padx=4)
        tk.Label(ar, text="난이도:").pack(side="left", padx=(8, 0))
        self.combo_min_diff = ttk.Combobox(ar, values=["1","2","3","4","5"], width=3, state="readonly"); self.combo_min_diff.set("2"); self.combo_min_diff.pack(side="left", padx=2)
        tk.Label(ar, text="~").pack(side="left")
        self.combo_max_diff = ttk.Combobox(ar, values=["1","2","3","4","5"], width=3, state="readonly"); self.combo_max_diff.set("4"); self.combo_max_diff.pack(side="left", padx=2)
        tk.Button(ar, text="자동 추출 → 장바구니", bg="#7c3aed", fg="white", font=("맑은 고딕", 9, "bold"), command=self._auto_select).pack(side="left", padx=12)

        # ── 장바구니 ──
        cf = tk.LabelFrame(self.frame, text=" 장바구니 (더블클릭=삭제) ", font=("맑은 고딕", 9, "bold"))
        cf.pack(fill="both", expand=True, padx=10, pady=8)
        self.list_cart = tk.Listbox(cf, height=6, selectmode=tk.EXTENDED)
        self.list_cart.pack(side="left", fill="both", expand=True, padx=4, pady=4)
        self.list_cart.bind("<Double-1>", self._cart_remove)

        bc = tk.Frame(cf); bc.pack(side="right", padx=8, fill="y")

        # ── 기본 옵션 ──
        fmt = tk.Frame(bc); fmt.pack(fill="x", pady=4)
        tk.Label(fmt, text="포맷:").pack(side="left")
        self.combo_format = ttk.Combobox(fmt, values=["Word + PDF", "Word만", "PDF만"], width=12, state="readonly")
        self.combo_format.current(0); self.combo_format.pack(side="left", padx=4)

        stf = tk.Frame(bc); stf.pack(fill="x", pady=4)
        tk.Label(stf, text="학생:").pack(side="left")
        self.combo_student = ttk.Combobox(stf, values=[], width=14, state="readonly")
        self.combo_student.pack(side="left", padx=4)

        shf = tk.Frame(bc); shf.pack(fill="x", pady=2)
        self.var_shuffle = tk.IntVar(value=0)
        tk.Checkbutton(shf, text="문항 셔플", variable=self.var_shuffle).pack(side="left")
        tk.Label(shf, text="세트:").pack(side="left", padx=(8, 0))
        self.combo_sets = ttk.Combobox(shf, values=["1","2","3"], width=3, state="readonly")
        self.combo_sets.set("1"); self.combo_sets.pack(side="left", padx=2)

        # ── v6: 출력 고급 설정 ──
        adv = tk.LabelFrame(bc, text=" ⚙️ 출력 고급 설정 ", font=("맑은 고딕", 8, "bold"), fg="#475569")
        adv.pack(fill="x", pady=6)

        # 시험지 제목
        title_f = tk.Frame(adv); title_f.pack(fill="x", padx=4, pady=2)
        tk.Label(title_f, text="제목:", font=("맑은 고딕", 8)).pack(side="left")
        self.ent_exam_title = tk.Entry(title_f, width=16, font=("맑은 고딕", 8))
        self.ent_exam_title.pack(side="left", padx=2)

        # 파일명
        fname_f = tk.Frame(adv); fname_f.pack(fill="x", padx=4, pady=2)
        tk.Label(fname_f, text="파일명:", font=("맑은 고딕", 8)).pack(side="left")
        self.ent_filename = tk.Entry(fname_f, width=16, font=("맑은 고딕", 8))
        self.ent_filename.pack(side="left", padx=2)

        # 배너
        banner_f = tk.Frame(adv); banner_f.pack(fill="x", padx=4, pady=2)
        tk.Button(banner_f, text="배너", font=("맑은 고딕", 8), command=self._select_banner).pack(side="left")
        tk.Button(banner_f, text="✕", font=("맑은 고딕", 8), width=2, command=self._remove_banner).pack(side="left", padx=2)
        self.lbl_banner = tk.Label(banner_f, text="(없음)", fg="#94a3b8", font=("맑은 고딕", 7))
        self.lbl_banner.pack(side="left", padx=2)

        # 빈 값 안내
        tk.Label(adv, text="(비워두면 기본값 자동 적용)", font=("맑은 고딕", 7), fg="#94a3b8").pack(padx=4, pady=1)

        # ── 생성 버튼 ──
        tk.Button(bc, text="시험지 생성", bg="#16a34a", fg="white", font=("맑은 고딕", 10, "bold"), width=20, command=lambda: self._generate("exam")).pack(pady=3)
        tk.Button(bc, text="워크북 생성", bg="#7c3aed", fg="white", font=("맑은 고딕", 10, "bold"), width=20, command=lambda: self._generate("workbook")).pack(pady=3)

        act = tk.Frame(bc); act.pack(fill="x", pady=4)
        tk.Button(act, text="선택 빼기", command=self._cart_remove).pack(side="left", expand=True, fill="x", padx=1)
        tk.Button(act, text="전체 비우기", command=self._cart_clear).pack(side="left", expand=True, fill="x", padx=1)

    # ── v6: 드래그 선택 (디바운싱) ──────────────

    def _on_drag_select(self, event):
        """마우스 드래그 시 연속 선택 — 새 행일 때만 처리 (Gemini 제안 디바운싱)"""
        item = self.tree.identify_row(event.y)
        if item and item != self._drag_last_row:
            self._drag_last_row = item
            self.tree.selection_add(item)

    def _on_drag_end(self, event):
        self._drag_last_row = None

    def _select_all_tree(self, event=None):
        """Ctrl+A — 포커스가 트리뷰일 때만 작동 (Entry 충돌 방지)"""
        focused = self.frame.focus_get()
        if focused != self.tree:
            return  # Entry 등 다른 위젯이면 기본 동작 유지
        all_items = []
        for item in self.tree.get_children():
            all_items.append(item)
            all_items.extend(self.tree.get_children(item))
        if all_items:
            self.tree.selection_set(all_items)
        return "break"

    # ── v6: 배너 선택 ─────────────────────────

    def _select_banner(self):
        path = filedialog.askopenfilename(
            title="배너 이미지 선택",
            filetypes=[("이미지 파일", "*.png *.jpg *.jpeg *.bmp *.gif")]
        )
        if path:
            self.banner_path = path
            self.lbl_banner.config(text=os.path.basename(path)[:18])

    def _remove_banner(self):
        self.banner_path = ""
        self.lbl_banner.config(text="(없음)")

    # ── 검색 ──────────────────────────────────

    def _reset_filters(self):
        self.ent_search.delete(0, tk.END)
        self.f_cat1.current(0); self.f_cat2.current(0); self.f_year.current(0)
        self.refresh()

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.preview_data.clear()

        passages = search_passages(self.cfg, search=self.ent_search.get().strip(),
                                   category1=self.f_cat1.get(), category2=self.f_cat2.get(), exam_year=self.f_year.get())
        for p in passages:
            pid = p["id"]; iid_p = f"P_{pid}"
            period = f"{p.get('school_year','')} {p.get('exam_year','')} {p.get('exam_month','')}".strip()
            tags = f"{p.get('publisher','')} {p.get('extra_tags','')}".strip()
            summary = (p.get("content","") or "")[:80].replace("\n"," ")
            self.tree.insert("", "end", iid=iid_p, values=(f"P-{pid}", "지문", period, tags, summary, "", p.get("created_at","")))
            self.preview_data[iid_p] = p.get("content","")

            for q in get_questions_for_passage(self.cfg, pid):
                iid_q = f"Q_{q['id']}_{pid}"
                diff = DIFFICULTY_LABELS.get(str(q.get("difficulty",3)), "중")
                ans_mark = f"[정답:{q['answer']}]" if q.get("answer") else ""
                q_summary = f"[{q['q_num']}번] {ans_mark} {(q.get('content','') or '')[:70].replace(chr(10),' ')}"
                self.tree.insert(iid_p, "end", iid=iid_q, values=(f"Q-{q['id']}", "문제", "", "", q_summary, diff, ""))
                self.preview_data[iid_q] = q.get("content","")

        students = list_students(self.cfg)
        self.combo_student["values"] = ["(선택안함)"] + [f"[{s['id']}] {s['name']}" for s in students]
        if self.combo_student["values"]:
            self.combo_student.current(0)

    def _on_motion(self, event):
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if iid and col == "#5" and self.preview_data.get(iid):
            self.tooltip.show(self.preview_data[iid], event.x_root, event.y_root)
        else:
            self.tooltip.hide()

    # ── 수정/삭제 ─────────────────────────────

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel: return messagebox.showwarning("알림", "삭제할 항목을 선택하세요.")
        if not messagebox.askyesno("경고", f"{len(sel)}개 항목을 영구 삭제하시겠습니까?"): return
        for iid in sel:
            if iid.startswith("P_"):
                pid = int(iid.split("_")[1]); delete_passage(self.cfg, pid); self.cart.pop(pid, None)
            elif iid.startswith("Q_"):
                parts = iid.split("_"); qid, pid = int(parts[1]), int(parts[2])
                delete_question(self.cfg, qid)
                if pid in self.cart and self.cart[pid]: self.cart[pid].discard(qid)
        self.refresh(); self._sync_cart_ui(); self.app.refresh_all()

    def _edit_selected(self):
        sel = self.tree.selection()
        if not sel: return
        iid = sel[0]
        if iid.startswith("Q_"): return messagebox.showinfo("안내", "지문 단위로 선택 후 수정하세요.")
        pid = int(iid.split("_")[1])
        p = get_passage(self.cfg, pid)
        qs = get_questions_for_passage(self.cfg, pid)
        if not p: return

        top = tk.Toplevel(self.app.root); top.title(f"수정 — 지문 P-{pid}"); top.geometry("900x850"); top.attributes("-topmost", True)
        tk.Label(top, text="지문", font=("맑은 고딕", 10, "bold")).pack(pady=4)
        txt_p = tk.Text(top, height=10, font=("Consolas", 10)); txt_p.pack(fill="x", padx=16); txt_p.insert("1.0", p["content"])
        tk.Label(top, text="문제 (사이에 빈 줄 필수)", font=("맑은 고딕", 10, "bold")).pack(pady=4)
        txt_q = tk.Text(top, height=10, font=("Consolas", 10)); txt_q.pack(fill="x", padx=16)
        for q in qs:
            num = f"{q['q_num']}. " if q['q_num'] != "-" else ""
            txt_q.insert(tk.END, f"{num}{q['content']}\n\n")
        tk.Label(top, text="정답 및 해설", font=("맑은 고딕", 10, "bold"), fg="#16a34a").pack(pady=4)
        txt_a = tk.Text(top, height=5, font=("Consolas", 10)); txt_a.pack(fill="x", padx=16); txt_a.insert("1.0", p.get("answer_text","") or "")

        def do_save():
            new_qs = []
            for qt in re.split(r"\n\s*\n(?=\d+\.\s)", txt_q.get("1.0", tk.END).strip()):
                if not qt.strip(): continue
                m = re.match(r"^(\d+)\.\s*(.*)", qt.strip(), re.DOTALL)
                new_qs.append({"q_num": m.group(1) if m else "-", "content": m.group(2) if m else qt.strip()})
            update_passage(self.cfg, pid, content=txt_p.get("1.0", tk.END).strip(),
                           answer_text=txt_a.get("1.0", tk.END).strip(), questions=new_qs)
            messagebox.showinfo("완료", "수정 반영 완료."); top.destroy(); self.refresh()

        tk.Button(top, text="수정 저장", bg="#2563eb", fg="white", font=("맑은 고딕", 11, "bold"), height=2, command=do_save).pack(fill="x", padx=16, pady=12)

    def _set_answer(self):
        sel = self.tree.selection()
        if not sel: return messagebox.showwarning("알림", "문항(Q)을 선택하세요.")
        for iid in sel:
            if not iid.startswith("Q_"): continue
            qid = int(iid.split("_")[1])
            q = get_question(self.cfg, qid)
            if not q: continue

            top = tk.Toplevel(self.app.root); top.title(f"정답/난이도 — Q-{qid}"); top.geometry("400x300"); top.attributes("-topmost", True)
            tk.Label(top, text=f"문항: {q['content'][:60]}...", wraplength=360).pack(pady=8, padx=16)

            f1 = tk.Frame(top); f1.pack(fill="x", padx=20, pady=4)
            tk.Label(f1, text="정답:").pack(side="left")
            ent_ans = tk.Entry(f1, width=8); ent_ans.pack(side="left", padx=4); ent_ans.insert(0, q.get("answer",""))
            tk.Label(f1, text="(①②③④⑤ 또는 A~E)", fg="#64748b").pack(side="left")

            f2 = tk.Frame(top); f2.pack(fill="x", padx=20, pady=4)
            tk.Label(f2, text="난이도:").pack(side="left")
            combo_d = ttk.Combobox(f2, values=["1-최하","2-하","3-중","4-상","5-최상"], width=8, state="readonly")
            combo_d.set(f"{q.get('difficulty',3)}-{DIFFICULTY_LABELS.get(str(q.get('difficulty',3)),'중')}"); combo_d.pack(side="left", padx=4)

            f3 = tk.Frame(top); f3.pack(fill="x", padx=20, pady=4)
            tk.Label(f3, text="해설:").pack(anchor="w")
            txt_exp = tk.Text(f3, height=4, font=("Consolas", 9)); txt_exp.pack(fill="x"); txt_exp.insert("1.0", q.get("explanation",""))

            def do_save(_qid=qid, _top=top):
                update_question_answer(self.cfg, _qid, ent_ans.get().strip(), txt_exp.get("1.0", tk.END).strip())
                update_question_difficulty(self.cfg, _qid, int(combo_d.get().split("-")[0]))
                _top.destroy(); self.refresh()

            tk.Button(top, text="저장", bg="#16a34a", fg="white", font=("맑은 고딕", 10, "bold"), command=do_save).pack(fill="x", padx=20, pady=12)
            break

    # ── 자동 출제 ─────────────────────────────

    def _auto_select(self):
        try: count = int(self.ent_auto_count.get())
        except ValueError: return messagebox.showwarning("오류", "문항수를 숫자로 입력하세요.")
        results = auto_select_questions(self.cfg, count=count, category1=self.f_cat1.get(), category2=self.f_cat2.get(),
                                        min_diff=int(self.combo_min_diff.get()), max_diff=int(self.combo_max_diff.get()))
        if not results: return messagebox.showinfo("결과", "조건에 맞는 문항이 없습니다.")
        for q in results:
            pid, qid = q["passage_id"], q["id"]
            if pid not in self.cart: self.cart[pid] = {qid}
            elif self.cart[pid]: self.cart[pid].add(qid)
        self._sync_cart_ui()
        messagebox.showinfo("추출 완료", f"{len(results)}문항 장바구니에 추가!")

    # ── 장바구니 ──────────────────────────────

    def _get_student_id(self):
        m = re.search(r"\[(\d+)\]", self.combo_student.get())
        return int(m.group(1)) if m else None

    def _sync_cart_ui(self):
        self.list_cart.delete(0, tk.END)
        sid = self._get_student_id()
        with db_conn(self.cfg) as conn:
            cur = conn.cursor()
            for pid, qids in self.cart.items():
                if not qids:
                    cnt = cur.execute("SELECT COUNT(*) FROM questions WHERE passage_id=?", (pid,)).fetchone()[0]
                    hist = ""
                    if sid:
                        q_rows = cur.execute("SELECT id FROM questions WHERE passage_id=?", (pid,)).fetchall()
                        total = sum(get_student_question_stats(self.cfg, sid, r[0])["count"] for r in q_rows)
                        if total: hist = f" [이력:{total}회]"
                    self.list_cart.insert("end", f"지문 전체 [ID:{pid}] ({cnt}문제){hist}")
                else:
                    for qid in qids:
                        qn_row = cur.execute("SELECT q_num FROM questions WHERE id=?", (qid,)).fetchone()
                        num = qn_row[0] if qn_row else "-"
                        hist = ""
                        if sid:
                            st = get_student_question_stats(self.cfg, sid, qid)
                            if st["count"]: hist = f" [이력:{st['count']}회, 최근:{st['last_date']}]"
                        self.list_cart.insert("end", f"개별 문제 [지문{pid} - {num}번 (ID:{qid})]{hist}")
        self.app.refresh_status()

    def _cart_add(self, event=None):
        sel = self.tree.selection()
        if not sel: return messagebox.showwarning("알림", "항목을 선택하세요.")
        added = 0
        for iid in sel:
            if iid.startswith("P_"):
                pid = int(iid.split("_")[1])
                if pid not in self.cart:
                    self.cart[pid] = set()
                    added += 1
            elif iid.startswith("Q_"):
                parts = iid.split("_"); qid, pid = int(parts[1]), int(parts[2])
                if pid not in self.cart:
                    self.cart[pid] = {qid}
                    added += 1
                elif self.cart[pid] and qid not in self.cart[pid]:
                    self.cart[pid].add(qid)
                    added += 1
        self._sync_cart_ui()
        # v6: 시각적 피드백
        if added:
            self.app.status_bar.set_text(f"✓ {added}개 항목 장바구니에 추가됨")

    def _cart_remove(self, event=None):
        for idx in reversed(list(self.list_cart.curselection())):
            text = self.list_cart.get(idx)
            m_full = re.search(r"\[ID:(\d+)\]", text)
            m_indiv = re.search(r"\(ID:(\d+)\)", text)
            if "지문 전체" in text and m_full:
                self.cart.pop(int(m_full.group(1)), None)
            elif m_indiv:
                qid = int(m_indiv.group(1))
                for pk, qs in list(self.cart.items()):
                    if qs and qid in qs: qs.discard(qid)
                    if qs is not None and not qs and pk in self.cart: del self.cart[pk]
        self._sync_cart_ui()

    def _cart_clear(self):
        self.cart.clear(); self._sync_cart_ui()

    # ── 생성 ──────────────────────────────────

    def _generate(self, mode):
        if not self.cart: return messagebox.showwarning("경고", "장바구니가 비어있습니다.")

        exam_data, all_qids = build_exam_data_from_cart(self.cfg, self.cart)
        is_wb = mode != "exam"
        prefix = "Workbook" if is_wb else "Test"

        sid = self._get_student_id()
        student_name = ""
        if sid:
            s = get_student(self.cfg, sid)
            student_name = s["name"] if s else ""

        num_sets = int(self.combo_sets.get())
        do_shuffle = self.var_shuffle.get()

        # v6: 사용자 지정 제목/파일명
        custom_title = self.ent_exam_title.get().strip()
        custom_filename = self.ent_filename.get().strip()

        # v6: 배너 (임시로 cfg 오버라이드)
        original_logo = self.cfg.get("last_logo", "")
        if self.banner_path:
            self.cfg["last_logo"] = self.banner_path

        messagebox.showinfo("생성 시작", "파일 생성을 시작합니다.\nPDF 변환 시 2~3초 소요될 수 있습니다.")

        try:
            if do_shuffle and num_sets > 1:
                sets = generate_multi_sets(all_qids, num_sets)
                for label, _ in sets:
                    result = create_exam_files(
                        self.cfg["last_dir"], prefix, exam_data, self.cfg,
                        is_workbook=is_wb, output_format=self.combo_format.get(),
                        student_name=student_name, set_label=label,
                        custom_title=custom_title, custom_filename=custom_filename,
                    )
            else:
                result = create_exam_files(
                    self.cfg["last_dir"], prefix, exam_data, self.cfg,
                    is_workbook=is_wb, output_format=self.combo_format.get(),
                    student_name=student_name,
                    custom_title=custom_title, custom_filename=custom_filename,
                )

            # 출제 기록 저장
            if sid and all_qids:
                create_exam_record(self.cfg, sid, all_qids, exam_title=f"{prefix}_{student_name}")

            messagebox.showinfo("완료", f"{result}\n\n저장: {self.cfg['last_dir']}")
            self._cart_clear()
            open_directory(self.cfg["last_dir"])

        except PermissionError:
            messagebox.showerror(
                "파일 접근 오류",
                "이전에 생성한 파일이 열려 있습니다.\n"
                "Word 또는 PDF 파일을 먼저 닫은 후 다시 시도해주세요."
            )
        except Exception as e:
            messagebox.showerror("생성 오류", f"오류가 발생했습니다:\n{e}")
        finally:
            # v6: 배너 cfg 복원
            self.cfg["last_logo"] = original_logo
