"""
상수 정의 모듈 — v7.0
"""

APP_NAME = "통합 문제은행 & 시험지 생성기 PRO v7.0"
APP_VERSION = "7.0.0"
SCHEMA_VERSION = 4

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

# v7.0: 예문 난이도 라벨
SENTENCE_DIFFICULTY = {1: "기본(중등)", 2: "핵심(고등내신)", 3: "심화(수능)"}
SENTENCE_DIFF_SHORT = {1: "기본", 2: "핵심", 3: "심화"}
SENTENCE_SOURCES = ["AI생성", "직접입력", "교과서본문"]

# v7.0: AI 프롬프트 — 예문 포함 JSON
AI_PROMPT_SENTENCES = """이 FactoryVoca 단어 시험지 사진에서 모든 단어를 추출하고,
각 단어에 대해 예문 1~2개를 함께 생성해주세요.

규칙:
1. 번호, 영어 단어, 한국어 뜻, 품사를 빠짐없이 포함
2. 빈칸에 쓴 답이 있으면 무시하고 원래 정답만 추출
3. 예문의 "target"에는 문장 내 실제 활용형을 적습니다 (experience→experienced)
4. 난이도(diff): 1=기본(중등), 2=핵심(고등내신), 3=심화(수능)

JSON 배열로만 출력하세요. 다른 설명은 필요 없습니다.
[
  {
    "no": 1, "section": "1과",
    "english": "experience", "korean": "경험", "pos": "n.",
    "sentences": [
      {"en": "I had a great experience abroad.", "ko": "나는 해외에서 좋은 경험을 했다.", "target": "experience", "diff": 1},
      {"en": "She experienced culture shock.", "ko": "그녀는 문화 충격을 경험했다.", "target": "experienced", "diff": 2}
    ]
  },
  ...
]"""

# v7.0: AI 프롬프트 — 단어만 (예문 없이, 기존 호환)
AI_PROMPT_WORDS_ONLY = """이 FactoryVoca 단어 시험지 사진에서 모든 단어를 추출해주세요.
번호, 영어 단어, 한국어 뜻, 품사를 빠짐없이 포함해야 합니다.
빈칸에 쓴 답이 있으면 무시하고 원래 정답만 추출하세요.

JSON 배열로만 출력하세요. 다른 설명은 필요 없습니다.
[
  {"no": 1, "english": "abandon", "korean": "버리다", "pos": "v."},
  ...
]"""

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
