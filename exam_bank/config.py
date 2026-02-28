"""
설정 관리 모듈 v5.0 — 경로, 폰트, JSON 설정 파일
"""

import os
import json
import platform
import subprocess

from exam_bank.constants import DEFAULT_CONFIG

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "exam_pro_config.json")


def setup_korean_font():
    """매트플롯립 한글 폰트 자동 설정 (matplotlib 있을 때만)."""
    try:
        import matplotlib.pyplot as plt
        import matplotlib.font_manager as fm
    except ImportError:
        return

    try:
        plt.style.use("seaborn-v0_8-whitegrid")
    except Exception:
        pass
    names = {f.name for f in fm.fontManager.ttflist}
    system = platform.system()
    if system == "Windows":
        for candidate in ("Malgun Gothic", "NanumGothic", "Gulim"):
            if candidate in names:
                plt.rcParams["font.family"] = candidate
                break
    elif system == "Darwin":
        plt.rcParams["font.family"] = "AppleGothic"
    else:
        plt.rcParams["font.family"] = "NanumGothic"
    plt.rcParams["axes.unicode_minus"] = False


def load_config():
    cfg = DEFAULT_CONFIG.copy()
    if not cfg["last_dir"]:
        cfg["last_dir"] = SCRIPT_DIR
    if not cfg["db_dir"]:
        cfg["db_dir"] = SCRIPT_DIR
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except (json.JSONDecodeError, OSError):
            pass
    return cfg


def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def get_db_path(cfg):
    folder = cfg.get("db_dir", SCRIPT_DIR)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "question_bank_v5.db")


def open_directory(path):
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])
