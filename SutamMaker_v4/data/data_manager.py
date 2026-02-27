"""
SutamMaker v4 — data/data_manager.py
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
File I/O helpers: history, draft (auto-save), question bank (JSON flat-file).
No Tkinter / docx / matplotlib dependencies.
"""

import os
import json
import datetime

from SutamMaker_v4.config import (
    QBANK_FILE,
    HISTORY_FILE,
    DRAFT_FILE,
)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 1. Generation history
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def add_history(entry):
    """Prepend *entry* to the history file (max 50 records)."""
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
    """Return the list of history records (newest first)."""
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
    """Persist an auto-save draft to disk."""
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
    """Load the latest draft, or return *None* if absent."""
    if not os.path.exists(DRAFT_FILE):
        return None
    try:
        with open(DRAFT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


def clear_draft():
    """Remove the draft file from disk."""
    if os.path.exists(DRAFT_FILE):
        try:
            os.remove(DRAFT_FILE)
        except OSError:
            pass


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 3. Question bank (JSON flat-file)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def save_to_qbank(questions, tag="", answers_map=None):
    """Append parsed questions to the JSON question bank.

    Parameters
    ----------
    questions : list[dict]
        Parsed question dicts (output of ``parse_exam_text``).
    tag : str
        User-supplied tag string.
    answers_map : dict | None
        ``{original_num: answer_text}`` mapping (optional).

    Returns
    -------
    int
        Number of questions saved.
    """
    bank = []
    if os.path.exists(QBANK_FILE):
        try:
            with open(QBANK_FILE, "r", encoding="utf-8") as f:
                bank = json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    ts = datetime.datetime.now().isoformat()
    if answers_map is None:
        answers_map = {}
    for q in questions:
        entry = {**q, "saved_at": ts, "tag": tag}
        q_num = q.get("num", "")
        ans_text = answers_map.get(q_num, "")
        entry["answer"] = ans_text
        bank.append(entry)
    with open(QBANK_FILE, "w", encoding="utf-8") as f:
        json.dump(bank, f, ensure_ascii=False, indent=2)
    return len(questions)


def load_qbank():
    """Return the full question bank as a list of dicts."""
    if not os.path.exists(QBANK_FILE):
        return []
    try:
        with open(QBANK_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def clear_qbank():
    """Delete the question bank file entirely."""
    if os.path.exists(QBANK_FILE):
        os.remove(QBANK_FILE)
