"""
SutamMaker v4 — data/data_manager.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
File I/O helpers: history, draft (auto-save), question bank (JSON flat-file).
No Tkinter / docx / matplotlib dependencies.
"""

import os
import json
import datetime

from config import QBANK_FILE, HISTORY_FILE, DRAFT_FILE


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 1. Generation history
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def add_history(entry):
    history = []
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                history = json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    history.insert(0, {**entry, "timestamp": datetime.datetime.now().isoformat()})
    history = history[:50]
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def load_history():
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 2. Auto-save draft
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def save_draft(exam_text, ans_text, title):
    try:
        with open(DRAFT_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "exam": exam_text,
                "answer": ans_text,
                "title": title,
                "saved_at": datetime.datetime.now().isoformat(),
            }, f, ensure_ascii=False)
    except OSError:
        pass


def load_draft():
    if not os.path.exists(DRAFT_FILE):
        return None
    try:
        with open(DRAFT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


def clear_draft():
    if os.path.exists(DRAFT_FILE):
        try:
            os.remove(DRAFT_FILE)
        except OSError:
            pass


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 3. Question bank (JSON flat-file)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def save_to_qbank(questions, tag="", answers_map=None):
    """Append parsed questions to the JSON question bank."""
    bank = load_qbank()
    ts = datetime.datetime.now().isoformat()
    if answers_map is None:
        answers_map = {}
    for q in questions:
        entry = {**q, "saved_at": ts, "tag": tag}
        q_num = q.get("num", "")
        entry["answer"] = answers_map.get(q_num, "")
        bank.append(entry)
    save_qbank_data(bank)
    return len(questions)


def load_qbank():
    if not os.path.exists(QBANK_FILE):
        return []
    try:
        with open(QBANK_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def save_qbank_data(bank):
    """Write the entire bank list to QBANK_FILE (used by delete/edit operations)."""
    with open(QBANK_FILE, "w", encoding="utf-8") as f:
        json.dump(bank, f, ensure_ascii=False, indent=2)


def clear_qbank():
    if os.path.exists(QBANK_FILE):
        os.remove(QBANK_FILE)
