"""
검증 유틸리티 — 선지 정규화, 정답 검증, 해시
"""

import hashlib
from exam_bank.constants import CIRCLE_NUMS, ALPHA_CHOICES, VALID_ANSWERS_NUM, VALID_ANSWERS_ALPHA


def normalize_choices(choices):
    result = list(choices or [])
    while len(result) < 5:
        result.append("")
    return result[:5]


def detect_choice_type(choices):
    for ch in (choices or []):
        if ch and any(c in ch for c in CIRCLE_NUMS):
            return "num"
        if ch and any(ch.strip().upper().startswith(a) for a in ALPHA_CHOICES):
            return "alpha"
    return "num"


def validate_answer(answer, choice_type="num"):
    if not answer or not answer.strip():
        return None
    answer = answer.strip()
    if choice_type == "alpha":
        return answer.upper() if answer.upper() in VALID_ANSWERS_ALPHA else None
    return answer if answer in VALID_ANSWERS_NUM else None


def compute_content_hash(passage, questions_text=""):
    raw = f"{passage.strip()}|{questions_text.strip()}"
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def check_answer(student_answer, correct_answer):
    if not student_answer or not correct_answer:
        return None
    s = student_answer.strip()
    c = correct_answer.strip()
    if not s or not c:
        return None
    return s == c
