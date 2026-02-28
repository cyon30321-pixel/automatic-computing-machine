"""
상수 정의 모듈 — v6.0
"""

APP_NAME = "통합 문제은행 & 시험지 생성기 PRO v6.0"
APP_VERSION = "6.0.0"
SCHEMA_VERSION = 3

CIRCLE_NUMS = "①②③④⑤"
ALPHA_CHOICES = ["A", "B", "C", "D", "E"]
VALID_ANSWERS_NUM = set(CIRCLE_NUMS)
VALID_ANSWERS_ALPHA = set(ALPHA_CHOICES)

SCHOOL_LEVELS = ["중학교", "고등학교"]
GRADES = ["1학년", "2학년", "3학년", "공통"]
EXAM_TYPES = ["내신", "모의고사", "부교재"]
EXAM_YEARS = ["2026년", "2025년", "2024년", "2023년"]
EXAM_MONTHS = ["3월", "4월", "6월", "9월", "11월"]
DIFFICULTY_LEVELS = ["1", "2", "3", "4", "5"]
DIFFICULTY_LABELS = {
    "1": "최하", "2": "하", "3": "중", "4": "상", "5": "최상"
}

# 단어장 카테고리
VOCAB_CATEGORIES = ["기본단어장", "시중단어장", "EBS", "부교재", "모의고사", "기타"]

DEFAULT_CONFIG = {
    "last_dir": "",
    "db_dir": "",
    "font_name": "맑은 고딕",
    "font_size": "10",
    "last_logo": "",
}


class Theme:
    PRIMARY      = "1B4F72"
    SECONDARY    = "2E86C1"
    ACCENT       = "E74C3C"
    LIGHT_BG     = "EBF5FB"
    PASSAGE_BG   = "F8F9FA"
    CONDITION_BG = "FFF9E6"
    TABLE_BORDER = "B0C4DE"
    HEADER_BG    = "D4E6F1"
    GRAY_TEXT    = "7F8C8D"
    DARK_TEXT    = "2C3E50"
    GREEN        = "27AE60"
    ANSWER_BG    = "D5F5E3"
