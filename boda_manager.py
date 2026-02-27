"""
BODA 매니저 PRO v2.0
─────────────────────
수업 중 화이트보드 캡처, 메모, 리포트를 한 곳에서 관리하는 도구.
"""

import os, sys, time, json, queue, ctypes, threading, winsound
from datetime import datetime, timedelta
import keyboard, pyperclip
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
from PIL import ImageGrab, ImageTk, Image

try:
    import pygetwindow as gw
except ImportError:
    print("필수 라이브러리 부족: pip install pygetwindow Pillow keyboard pyperclip")
    sys.exit(1)

try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
    try: ctypes.windll.user32.SetProcessDPIAware()
    except: pass

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "boda_config.json")
DEFAULT_CONFIG = {
    "base_path": r"C:\Users\cyon3\OneDrive\캡처\2026\2월",
    "window_keywords": ["BODA", "상상", "보다"],
    "crop_left_pct": 0.31, "crop_top_px": 85, "crop_right_px": 115, "crop_bottom_px": 70,
    "capture_delay": 0.35, "autosave_interval_sec": 120, "recent_students_max": 10,
    "hotkeys": {"capture": ["f8", "f9"], "change_student": ["f10", "f11"]},
}

COLORS = {"bg":"#f5f7fa","card":"#ffffff","primary":"#3b82f6","primary_dark":"#2563eb",
    "success":"#10b981","warning":"#f59e0b","danger":"#ef4444","text":"#1e293b",
    "text_light":"#64748b","border":"#e2e8f0","accent":"#8b5cf6"}

def beep(freq, ms):
    try: winsound.Beep(freq, ms)
    except: pass
def beep_async(freq, ms): threading.Thread(target=beep, args=(freq, ms), daemon=True).start()
def sound_capture(): beep_async(1200, 80)
def sound_success(): beep_async(1000, 100)
def sound_switch(): beep_async(1500, 80); time.sleep(0.1); beep_async(1800, 80)
def sound_error(): beep_async(400, 300)

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f: saved = json.load(f)
            merged = {**DEFAULT_CONFIG, **saved}
            merged["hotkeys"] = {**DEFAULT_CONFIG["hotkeys"], **saved.get("hotkeys", {})}
            return merged
        except: pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(cfg, f, ensure_ascii=False, indent=2)

class BodaManager:
    def __init__(self):
        self.cfg = load_config()
        self.current_student = ""
        self.session_start = None
        self.capture_count = 0
        self.recent_students = []
        self.msg_queue = queue.Queue()
        self.last_thumbnail = None
        self._load_recent_students()
        self._minimize_console()
        self._build_ui()
        self._register_hotkeys()
        threading.Thread(target=self._console_listener, daemon=True).start()

    def _build_ui(self):
        self.root = tk.Tk()
        self.root.title("BODA 매니저 PRO v2")
        self.root.configure(bg=COLORS["bg"])
        self.root.attributes("-topmost", True)
        self.root.geometry("360x640")
        self.root.minsize(340, 580)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        style = ttk.Style(); style.theme_use("clam")

        header = tk.Frame(self.root, bg=COLORS["primary"], height=56)
        header.pack(fill="x"); header.pack_propagate(False)
        tk.Label(header, text="BODA 매니저 PRO", font=("맑은 고딕", 13, "bold"), bg=COLORS["primary"], fg="white").pack(side="left", padx=14, pady=12)
        self.lbl_timer = tk.Label(header, text="00:00:00", font=("Consolas", 12), bg=COLORS["primary"], fg="#bfdbfe")
        self.lbl_timer.pack(side="right", padx=14)

        student_card = tk.Frame(self.root, bg=COLORS["card"], bd=0, highlightthickness=1, highlightbackground=COLORS["border"])
        student_card.pack(fill="x", padx=12, pady=(12, 6))
        row = tk.Frame(student_card, bg=COLORS["card"]); row.pack(fill="x", padx=12, pady=10)
        self.lbl_student = tk.Label(row, text="학생 미선택", font=("맑은 고딕", 14, "bold"), bg=COLORS["card"], fg=COLORS["text"])
        self.lbl_student.pack(side="left")
        self.lbl_captures = tk.Label(row, text="캡처 0장", font=("맑은 고딕", 9), bg=COLORS["card"], fg=COLORS["text_light"])
        self.lbl_captures.pack(side="right")

        btn_frame = tk.Frame(self.root, bg=COLORS["bg"]); btn_frame.pack(fill="x", padx=12, pady=6)
        self.btn_capture = tk.Button(btn_frame, text="  캡처  [F8/F9]", font=("맑은 고딕", 10, "bold"), bg=COLORS["success"], fg="white", bd=0, padx=16, pady=10, command=lambda: self.msg_queue.put("cap"))
        self.btn_capture.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self.btn_switch = tk.Button(btn_frame, text="  학생 변경  [F10/F11]", font=("맑은 고딕", 10, "bold"), bg=COLORS["primary"], fg="white", bd=0, padx=16, pady=10, command=lambda: self.msg_queue.put("name"))
        self.btn_switch.pack(side="left", expand=True, fill="x", padx=(4, 0))

        thumb_card = tk.Frame(self.root, bg=COLORS["card"], bd=0, highlightthickness=1, highlightbackground=COLORS["border"])
        thumb_card.pack(fill="x", padx=12, pady=6)
        tk.Label(thumb_card, text="마지막 캡처", font=("맑은 고딕", 8), bg=COLORS["card"], fg=COLORS["text_light"]).pack(anchor="w", padx=10, pady=(6, 0))
        self.lbl_thumb = tk.Label(thumb_card, text="캡처 없음", font=("맑은 고딕", 9), bg=COLORS["border"], fg=COLORS["text_light"], width=40, height=5)
        self.lbl_thumb.pack(padx=10, pady=(4, 10), fill="x")

        memo_card = tk.Frame(self.root, bg=COLORS["card"], bd=0, highlightthickness=1, highlightbackground=COLORS["border"])
        memo_card.pack(fill="both", expand=True, padx=12, pady=6)
        memo_header = tk.Frame(memo_card, bg=COLORS["card"]); memo_header.pack(fill="x", padx=10, pady=(8, 0))
        tk.Label(memo_header, text="수업 메모", font=("맑은 고딕", 9, "bold"), bg=COLORS["card"], fg=COLORS["text"]).pack(side="left")
        self.lbl_autosave = tk.Label(memo_header, text="", font=("맑은 고딕", 7), bg=COLORS["card"], fg=COLORS["success"])
        self.lbl_autosave.pack(side="right")
        self.memo_box = tk.Text(memo_card, font=("맑은 고딕", 10), wrap="word", bd=0, bg="#f8fafc", fg=COLORS["text"], padx=8, pady=8, relief="flat", highlightthickness=1, highlightbackground=COLORS["border"])
        self.memo_box.pack(fill="both", expand=True, padx=10, pady=(4, 4))

        status_bar = tk.Frame(self.root, bg=COLORS["border"], height=28)
        status_bar.pack(fill="x", side="bottom"); status_bar.pack_propagate(False)
        self.lbl_status = tk.Label(status_bar, text="시작하려면 학생 이름을 입력하세요", font=("맑은 고딕", 8), bg=COLORS["border"], fg=COLORS["text_light"])
        self.lbl_status.pack(side="left", padx=10)

    def run(self):
        self.root.after(100, self._prompt_first_student)
        self.root.after(200, self._poll_queue)
        self.root.after(1000, self._tick_timer)
        self.root.after(self.cfg["autosave_interval_sec"] * 1000, self._autosave_memo)
        self.root.mainloop()

    def _prompt_first_student(self):
        name = simpledialog.askstring("시작", "첫 학생 이름:", parent=self.root)
        if not name: self.root.destroy(); return
        self._set_student(name)

    def _set_student(self, name):
        self.current_student = name.strip()
        self.session_start = datetime.now()
        self.capture_count = 0
        self._add_recent(name.strip())
        self.lbl_student.config(text=self.current_student)
        self.lbl_captures.config(text="캡처 0장")
        self._set_status(f"'{self.current_student}' 수업 시작")
        sound_success()

    def _add_recent(self, name):
        if name in self.recent_students: self.recent_students.remove(name)
        self.recent_students.insert(0, name)
        self.recent_students = self.recent_students[:self.cfg["recent_students_max"]]
        self._save_recent_students()

    def _load_recent_students(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".boda_recent.json")
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f: self.recent_students = json.load(f)
            except: self.recent_students = []

    def _save_recent_students(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".boda_recent.json")
        with open(path, "w", encoding="utf-8") as f: json.dump(self.recent_students, f, ensure_ascii=False)

    def _tick_timer(self):
        if self.session_start:
            elapsed = datetime.now() - self.session_start
            total_sec = int(elapsed.total_seconds())
            h, r = divmod(total_sec, 3600); m, s = divmod(r, 60)
            self.lbl_timer.config(text=f"{h:02d}:{m:02d}:{s:02d}")
        self.root.after(1000, self._tick_timer)

    def _capture(self):
        if not self.current_student: self._set_status("학생을 먼저 선택하세요"); sound_error(); return
        self._set_status("캡처 중..."); self.btn_capture.config(state="disabled")
        try:
            self.root.withdraw(); time.sleep(self.cfg["capture_delay"])
            wins = [w for w in gw.getAllWindows() if any(k in w.title for k in self.cfg["window_keywords"])]
            if not wins: self._set_status("BODA 창을 찾을 수 없습니다"); sound_error(); return
            win = wins[0]
            if win.isMinimized: win.restore()
            try: win.activate()
            except: pass
            time.sleep(0.25)
            left_off = int(win.width * self.cfg["crop_left_pct"])
            bbox = (win.left + left_off, win.top + self.cfg["crop_top_px"], win.right - self.cfg["crop_right_px"], win.bottom - self.cfg["crop_bottom_px"])
            img = ImageGrab.grab(bbox=bbox, all_screens=True)
            now = datetime.now()
            folder = os.path.join(self.cfg["base_path"], now.strftime("%Y-%m-%d"), self.current_student)
            os.makedirs(folder, exist_ok=True)
            filename = f"{self.current_student}_{now.strftime('%H%M%S')}.png"
            img.save(os.path.join(folder, filename))
            self.capture_count += 1
            self.lbl_captures.config(text=f"캡처 {self.capture_count}장")
            self._set_status(f"캡처 저장 완료 ({filename})"); sound_capture()
        except Exception as e: self._set_status(f"캡처 실패: {e}"); sound_error()
        finally: self.root.deiconify(); self.btn_capture.config(state="normal")

    def _save_memo(self):
        content = self.memo_box.get("1.0", tk.END).strip()
        if not self.current_student or not content: return
        now = datetime.now()
        folder = os.path.join(self.cfg["base_path"], now.strftime("%Y-%m-%d"), self.current_student)
        os.makedirs(folder, exist_ok=True)
        filename = f"{self.current_student}_수업메모_{now.strftime('%Y%m%d')}.txt"
        full_path = os.path.join(folder, filename)
        mode = "a" if os.path.exists(full_path) else "w"
        with open(full_path, mode, encoding="utf-8") as f:
            if mode == "w": f.write(f"=== {self.current_student} 수업 메모 ({now.strftime('%Y-%m-%d')}) ===\n\n")
            f.write(f"[{now.strftime('%H:%M')}]\n{content}\n\n")
        self.memo_box.delete("1.0", tk.END)

    def _autosave_memo(self):
        content = self.memo_box.get("1.0", tk.END).strip()
        if content and self.current_student:
            self._save_memo()
            self.lbl_autosave.config(text=f"자동저장 {datetime.now().strftime('%H:%M')}")
        self.root.after(self.cfg["autosave_interval_sec"] * 1000, self._autosave_memo)

    def _poll_queue(self):
        try:
            while not self.msg_queue.empty():
                task = self.msg_queue.get_nowait()
                if task == "cap": self._capture()
                elif task == "name":
                    self._save_memo(); sound_switch()
                    new_name = simpledialog.askstring("학생 변경", "새 학생 이름:", parent=self.root)
                    if new_name: self._set_student(new_name)
        except: pass
        self.root.after(100, self._poll_queue)

    def _register_hotkeys(self):
        for key in self.cfg["hotkeys"]["capture"]: keyboard.add_hotkey(key, lambda: self.msg_queue.put("cap"))
        for key in self.cfg["hotkeys"]["change_student"]: keyboard.add_hotkey(key, lambda: self.msg_queue.put("name"))

    def _console_listener(self):
        while True:
            try:
                cmd = input().strip().lower()
                if cmd == "c": self.msg_queue.put("cap")
                elif cmd == "n": self.msg_queue.put("name")
                elif cmd == "q": self.msg_queue.put("quit")
            except EOFError: break

    def _set_status(self, text): self.lbl_status.config(text=text)

    @staticmethod
    def _minimize_console():
        try:
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd: ctypes.windll.user32.ShowWindow(hwnd, 6)
        except: pass

    def _on_close(self):
        self._save_memo(); self.root.destroy()

if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    app = BodaManager(); app.run()
