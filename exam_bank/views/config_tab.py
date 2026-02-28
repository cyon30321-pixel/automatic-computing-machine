"""
설정 탭 — DB 경로, 폰트, 백업/복원
"""

import os
import json
import tkinter as tk
from tkinter import messagebox, filedialog, ttk

from exam_bank.config import save_config, get_db_path, open_directory
from exam_bank.models.database import db_stats
from exam_bank.services.backup import (
    backup_db, backup_db_zip, restore_db, list_backups,
    export_json, import_json,
)


class ConfigTab:
    def __init__(self, parent, cfg, app):
        self.cfg = cfg
        self.app = app
        self.frame = ttk.Frame(parent)
        self._build()

    def _build(self):
        main = tk.Frame(self.frame)
        main.pack(fill="both", expand=True, padx=20, pady=15)

        # DB 정보
        info = tk.LabelFrame(main, text=" DB 정보 ", font=("맑은 고딕", 10, "bold"), pady=8)
        info.pack(fill="x", pady=(0, 8))
        self.lbl_db_path = tk.Label(info, text="", anchor="w")
        self.lbl_db_path.pack(fill="x", padx=10)
        self.lbl_stats = tk.Label(info, text="", anchor="w", fg="#475569")
        self.lbl_stats.pack(fill="x", padx=10)

        btns = tk.Frame(info)
        btns.pack(padx=10, pady=8, anchor="w")
        tk.Button(btns, text="DB 폴더 변경", command=self._change_db_dir).pack(side="left", padx=4)
        tk.Button(btns, text="DB 폴더 열기", command=lambda: open_directory(self.cfg.get("db_dir", ""))).pack(side="left", padx=4)

        # 백업
        bk = tk.LabelFrame(main, text=" 백업 / 복원 ", font=("맑은 고딕", 10, "bold"), pady=8)
        bk.pack(fill="x", pady=(0, 8))
        br = tk.Frame(bk)
        br.pack(padx=10, pady=4, anchor="w")
        tk.Button(br, text="DB 백업", bg="#2563eb", fg="white", command=self._backup).pack(side="left", padx=4)
        tk.Button(br, text="ZIP 백업", bg="#7c3aed", fg="white", command=self._backup_zip).pack(side="left", padx=4)
        tk.Button(br, text="복원", bg="#dc2626", fg="white", command=self._restore).pack(side="left", padx=4)

        # JSON Import/Export
        je = tk.LabelFrame(main, text=" JSON Import/Export ", font=("맑은 고딕", 10, "bold"), pady=8)
        je.pack(fill="x", pady=(0, 8))
        jr = tk.Frame(je)
        jr.pack(padx=10, pady=4, anchor="w")
        tk.Button(jr, text="JSON 내보내기", bg="#16a34a", fg="white", command=self._export_json).pack(side="left", padx=4)
        tk.Button(jr, text="JSON 가져오기", bg="#ea580c", fg="white", command=self._import_json).pack(side="left", padx=4)

        self.refresh_stats()

    def refresh_stats(self):
        self.lbl_db_path.config(text=f"DB 경로: {get_db_path(self.cfg)}")
        try:
            s = db_stats(self.cfg)
            self.lbl_stats.config(text=f"지문 {s['passages']}개 | 문제 {s['questions']}개 | 학생 {s['students']}명 | 시험 {s['exams']}건")
        except Exception:
            self.lbl_stats.config(text="DB 연결 실패")

    def _change_db_dir(self):
        d = filedialog.askdirectory(title="DB 폴더 선택")
        if d:
            self.cfg["db_dir"] = d
            save_config(self.cfg)
            from exam_bank.models.database import init_db
            init_db(self.cfg)
            self.refresh_stats()
            self.app.refresh_all()
            messagebox.showinfo("완료", f"DB 경로 변경: {d}")

    def _backup(self):
        try:
            path = backup_db(self.cfg)
            messagebox.showinfo("백업 완료", f"저장: {path}")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _backup_zip(self):
        try:
            path = backup_db_zip(self.cfg)
            messagebox.showinfo("ZIP 백업 완료", f"저장: {path}")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _restore(self):
        path = filedialog.askopenfilename(title="복원 파일 선택", filetypes=[("DB/ZIP", "*.db *.zip")])
        if not path: return
        if not messagebox.askyesno("복원 확인", "현재 DB를 덮어씁니다. 계속하시겠습니까?"): return
        try:
            restore_db(self.cfg, path)
            self.app.refresh_all()
            messagebox.showinfo("완료", "복원 완료!")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _export_json(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if not path: return
        try:
            data = export_json(self.cfg)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("완료", f"{len(data)}개 지문 내보내기 완료")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _import_json(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not path: return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            count = import_json(self.cfg, data)
            self.app.refresh_all()
            messagebox.showinfo("완료", f"{count}개 지문 가져오기 완료")
        except Exception as e:
            messagebox.showerror("오류", str(e))
