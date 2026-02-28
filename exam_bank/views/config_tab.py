"""
설정 및 백업 탭
"""

import os
import tkinter as tk
from tkinter import messagebox, filedialog, ttk

from exam_bank.models.database import db_stats, init_db
from exam_bank.services.backup import backup_db, backup_db_zip, restore_db, list_backups
from exam_bank.config import save_config, open_directory, SCRIPT_DIR


class ConfigTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self._build()

    def _build(self):
        main = tk.Frame(self.frame)
        main.pack(fill="both", expand=True, padx=30, pady=20)

        pf = tk.LabelFrame(main, text=" 경로 설정 ", font=("맑은 고딕", 10, "bold"))
        pf.pack(fill="x", pady=(0, 12))
        r0 = tk.Frame(pf); r0.pack(fill="x", padx=10, pady=6)
        tk.Label(r0, text="출력 폴더:", width=18, anchor="w").pack(side="left")
        self.lbl_out = tk.Label(r0, text=self.cfg["last_dir"], fg="#2563eb"); self.lbl_out.pack(side="left", padx=8)
        tk.Button(r0, text="변경", command=self._change_out_dir).pack(side="right", padx=4)
        tk.Button(r0, text="열기", command=lambda: open_directory(self.cfg["last_dir"])).pack(side="right", padx=4)

        r1 = tk.Frame(pf); r1.pack(fill="x", padx=10, pady=6)
        tk.Label(r1, text="DB 폴더:", width=18, anchor="w").pack(side="left")
        self.lbl_db = tk.Label(r1, text=self.cfg.get("db_dir", SCRIPT_DIR), fg="#16a34a"); self.lbl_db.pack(side="left", padx=8)
        tk.Button(r1, text="변경", command=self._change_db_dir).pack(side="right", padx=4)
        tk.Button(r1, text="열기", command=lambda: open_directory(self.cfg.get("db_dir", SCRIPT_DIR))).pack(side="right", padx=4)

        bf = tk.LabelFrame(main, text=" DB 백업 및 복원 ", font=("맑은 고딕", 10, "bold"))
        bf.pack(fill="x", pady=12)
        br = tk.Frame(bf); br.pack(fill="x", padx=10, pady=10)
        tk.Button(br, text="DB 백업 (.db)", bg="#2563eb", fg="white", font=("맑은 고딕", 10, "bold"), width=16, command=self._backup_db).pack(side="left", padx=4)
        tk.Button(br, text="ZIP 백업", bg="#0891b2", fg="white", font=("맑은 고딕", 10, "bold"), width=16, command=self._backup_zip).pack(side="left", padx=4)
        tk.Button(br, text="백업에서 복원", bg="#ea580c", fg="white", font=("맑은 고딕", 10, "bold"), width=16, command=self._restore_db).pack(side="left", padx=4)
        self.lbl_backup_info = tk.Label(bf, text="", font=("맑은 고딕", 9), fg="#64748b")
        self.lbl_backup_info.pack(padx=10, pady=(0, 8))
        self._update_backup_info()

        sf = tk.LabelFrame(main, text=" DB 현황 ", font=("맑은 고딕", 10, "bold"))
        sf.pack(fill="x", pady=12)
        self.lbl_stats = tk.Label(sf, text="", font=("맑은 고딕", 11), pady=10); self.lbl_stats.pack()
        self.refresh_stats()

    def refresh_stats(self):
        try:
            s = db_stats(self.cfg)
            self.lbl_stats.config(text=f"지문: {s['passages']}개   문제: {s['questions']}개   학생: {s['students']}명   시험: {s['exams']}건")
        except: self.lbl_stats.config(text="DB 연결 오류")

    def _change_out_dir(self):
        d = filedialog.askdirectory()
        if d: self.cfg["last_dir"] = d; self.lbl_out.config(text=d); save_config(self.cfg)

    def _change_db_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.cfg["db_dir"] = d; self.lbl_db.config(text=d); save_config(self.cfg)
            init_db(self.cfg); self.app.refresh_all()
            messagebox.showinfo("완료", "DB 경로 변경 완료.")

    def _backup_db(self):
        try: path = backup_db(self.cfg); messagebox.showinfo("백업 완료", f"백업 파일:\n{path}"); self._update_backup_info()
        except Exception as e: messagebox.showerror("오류", f"백업 실패: {e}")

    def _backup_zip(self):
        try: path = backup_db_zip(self.cfg); messagebox.showinfo("ZIP 백업 완료", f"백업 파일:\n{path}"); self._update_backup_info()
        except Exception as e: messagebox.showerror("오류", f"백업 실패: {e}")

    def _restore_db(self):
        src = filedialog.askopenfilename(title="복원할 백업 파일 선택",
            filetypes=[("DB/ZIP 파일", "*.db *.zip")], initialdir=self.cfg.get("db_dir", SCRIPT_DIR))
        if not src: return
        if not messagebox.askyesno("경고", "현재 DB를 선택한 백업으로 덮어씁니다.\n계속하시겠습니까?"): return
        try:
            restore_db(self.cfg, src); init_db(self.cfg); self.app.refresh_all()
            messagebox.showinfo("복원 완료", "DB 복원이 완료되었습니다.")
        except Exception as e: messagebox.showerror("오류", f"복원 실패: {e}")

    def _update_backup_info(self):
        backups = list_backups(self.cfg)
        self.lbl_backup_info.config(text=f"최근 백업: {backups[0]['name']}" if backups else "백업 파일 없음")
