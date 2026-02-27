"""
SutamMaker v4 — config.py
━━━━━━━━━━━━━━━━━━━━━━━━━
All global constants, file-path helpers, and persistent config (JSON) live here.
No Tkinter / docx / matplotlib dependencies allowed.
"""

import os
import json
import hashlib
import shutil
import tempfile
import datetime

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 0. File-path constants
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
try:
    os.chdir(SCRIPT_DIR)
except OSError:
    pass

CONFIG_FILE = os.path.join(SCRIPT_DIR, "sutam_maker_config.json")
QBANK_FILE = os.path.join(SCRIPT_DIR, "sutam_question_bank.json")
HISTORY_FILE = os.path.join(SCRIPT_DIR, "sutam_history.json")
DRAFT_FILE = os.path.join(SCRIPT_DIR, ".sutam_draft.json")
MAX_SAFE_PATH = 240

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 1. UI / Domain constants
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CIRCLE_NUMS = "①②③④⑤"

SUBJECT_LIST = [
    "경제", "사회·문화", "생활과 윤리", "윤리와 사상",
    "정치와 법", "한국지리", "세계지리", "동아시아사",
    "세계사", "한국사",
]

DIFFICULTY_LIST = [
    "하~중 (교과서 기본 개념)",
    "중상 (개념 융합 + 함정 선지)",
    "최상/킬러 (수능 오답률 TOP5 모방)",
    "혼합 (하~중 40%% + 중상 40%% + 킬러 20%%)",
]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 2. AI Master Prompt Template
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
AI_MASTER_PROMPT_TEMPLATE = """
============================================================
0) ROLE & GOAL
============================================================
너는 한국의 수능/평가원 모의고사 및 대치동 자사고 내신을 20년간 출제해 온 '사회탐구 및 경제 영역 최고 권위자'다.
사용자의 요청(특정 단원, 난이도, 문항 수 등)에 따라 실전 내신/수능 대비용 시험지를 생성한다.
생성된 텍스트는 파이썬(Python) 파싱 프로그램과 100% 호환되어야 하므로, 출력 형식과 태그 규칙을 엄격하게 준수해야 한다.
============================================================
1) PHASE 1 — WEB SEARCH & DATA SOURCING (트렌드 분석 및 모방)
============================================================
문항 생성 전, 반드시 최신 수능/평가원 기출문제 및 주요 학군 내신 트렌드를 웹 검색으로 수집하여 반영한다.
- Lower-Mid (하~중): 교과서 기본 개념 확인 및 단순 표 해석
- Upper-Mid (중상): 두 가지 이상의 개념 융합, 일반적인 함정 선지 포함
- High/Killer (최상): 최근 3개년 수능 오답률 1~5위 문항 모방. (경제: 복잡한 계산 및 예외적 그래프 이동 / 사회문화: 계층 이동, 빈곤율 등 표 풀이 / 생활과윤리: 사상가 3자 벤다이어그램 비교 등)
============================================================
2) PHASE 2 — 사탐/경제 문항 생성 절대 규칙 (PYTHON 연동)
============================================================
파이썬 프로그램이 문항을 인식하고 워드 파일 및 그래프를 생성할 수 있도록 아래 규칙을 무조건 따른다.
[1. 문항 기본 포맷]
- 모든 문제 번호는 "1.", "2.", "03." 형식으로 시작한다. (Q1 등 사용 금지)
- 선지는 반드시 원문자 "①, ②, ③, ④, ⑤"를 사용한다.
[2. 사탐 전용 <보기> 박스 규칙]
- 수능형 ㄱ, ㄴ, ㄷ, ㄹ 선지가 필요한 경우, 발문 바로 아래에 반드시 '<보기>' 라는 태그를 넣는다.
- 파이썬이 이를 인식하여 워드에 예쁜 박스(Table)를 그릴 것이다.
(예시 형식)
<보기>
ㄱ. X재의 가격은 상승한다.
ㄴ. Y재의 거래량은 감소한다.
[3. 경제 그래프 자동 생성 태그 (GRAPH TAG) - 매우 중요]
경제 그래프가 필요한 문항은 파이썬이 Matplotlib으로 자동 작도할 수 있도록 발문과 <보기> 사이에 반드시 지정된 태그를 삽입한다.
- 형태 A (듀얼 개념 그래프 - X재/Y재 시장 등):
  [DUAL_GRAPH | Left: Shift_S=Right | Right: Shift_D=Left]
  (방향은 Right, Left만 사용. S는 공급, D는 수요)
- 형태 B (단일 수치 그래프 - 보조금/세금 등):
  [GRAPH: P=[200,300,400] | Qd=[120,110,100] | Qs=[80,90,100] | Shift_S=-200]
  (P, Qd, Qs의 배열 길이는 동일해야 함. Shift_S 또는 Shift_D로 곡선 이동 명시)
- 형태 C (PPF 생산가능곡선):
  [PPF | PPF_SHIFT=OUT]  (OUT=바깥 이동, IN=안쪽 이동)
- 형태 D (AS-AD 총수요-총공급):
  [ASAD | Shift_AD=Right | Shift_SRAS=Left]
- 형태 E (로렌츠 곡선):
  [LORENZ | GINI=0.4 | COMPARE_GINI=0.3]
- 형태 F (필립스 곡선):
  [PHILLIPS | PHILLIPS_SHIFT=UP]
- 형태 G (조세 부과):
  [TAX]
- 형태 H (관세):
  [TARIFF]
- 형태 I (외부효과):
  [EXTERNALITY]
- 형태 J (무차별곡선/예산선):
  [INDIFFERENCE]
[4-1. 이미지 삽입 태그]
사용자가 제공한 로컬 이미지를 문항에 삽입하려면, 발문 내에 [IMAGE:파일경로] 태그를 사용한다.
(예시) [IMAGE:C:/자료/그림1.png]
[4. 표(Table) 데이터 출력 규칙]
사탐/경제에서 표 데이터가 필요한 경우, 텍스트로 풀어쓰지 말고 반드시 마크다운 형식(| 항목 | 값 |)의 표로 깔끔하게 작성할 것.
[5. 제시문 및 사료 박스 규칙]
일반적인 대화문(민수: ~, 영희: ~)이나 사료, 긴 글 지문이 있을 경우, 텍스트 앞뒤로 반드시 <제시문> 과 </제시문> 태그를 감싸줄 것.
============================================================
3) PHASE 3 — 매력적인 오답(함정) 설계 Blueprint
============================================================
단순히 틀린 말을 적는 것이 아니라, 수험생이 빠지기 쉬운 논리적 함정을 설계한다.
- 경제: '곡선 상의 이동'과 '곡선 자체의 이동' 혼동 유발 / 기회비용 절대우위와 비교우위 수치 혼동 / 탄력성(탄력적vs비탄력적)에 따른 총수입 변화 방향 반대로 서술.
- 일반 사탐: A학자의 주장을 B학자의 주장인 것처럼 교묘하게 서술 / '모든', '항상', '반드시' 등 극단적 한정사를 이용한 오답 / 표 분석에서 변화율(%%)과 변화량(%%p) 혼동 유발.
============================================================
4) OUTPUT FORMAT (반드시 2개 블록으로 분리 출력)
============================================================
출력은 파이썬이 완벽히 파싱할 수 있도록 아래 구조를 엄격히 지킨다.
[출력 1 — 학생용 시험지]
(어떠한 인사말이나 메모도 없이 바로 1번 문제부터 시작. 지문, 태그, <보기>, ①~⑤ 선지 이외의 텍스트 절대 금지)
[출력 2 — 교사용 정답 및 해설]
- 상단에 반드시 "1번 정답: ③", "2번 정답: ⑤" 형태로 한 줄에 하나씩 정답만 나열.
- 그 아래에 각 문항별 [해설] 및 [오답 피하기(오답 선지가 틀린 이유)]를 상세히 서술.
- 경제/사문 표 풀이 문항의 경우, [Step-by-Step 계산 논리]를 수식과 함께 제시.
============================================================
[사용자 요청]
- 과목: {subject}
- 난이도: {difficulty}
- 문항 수: {num_questions}문항
- 범위/단원: {scope}
- 추가 지시: {extra}
============================================================
요청에 따라 즉시 기출 데이터를 웹 검색하고, 위의 완벽한 포맷으로 시험지를 생성하라.
""".strip()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 3. Persistent config (load / save)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            pass
    return {
        "last_dir": SCRIPT_DIR, "last_logo": "", "font_name": "맑은 고딕", "font_size": "10",
        "db_root": os.path.join(SCRIPT_DIR, "db"),
        "academy_name": "", "exam_name": "", "contact_info": "", "copyright_text": "",
        "show_page_number": False, "show_date": False,
    }


def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 4. Path helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _safe_basename(folder, stem, ext):
    full = os.path.join(folder, stem + ext)
    if len(full) <= MAX_SAFE_PATH:
        return stem
    h = hashlib.md5(stem.encode("utf-8")).hexdigest()[:6]
    cut = max(1, MAX_SAFE_PATH - (len(folder) + 1 + len(ext) + 8))
    return f"{stem[:cut]}_{h}"


def build_paths(target_dir, base_filename, prefix=""):
    stem = _safe_basename(target_dir, f"{prefix}{base_filename}", ".docx")
    return os.path.join(target_dir, stem + ".docx"), os.path.join(target_dir, stem + ".pdf")


def convert_docx_pairs_to_pdf(pairs):
    """Convert a list of (docx_path, pdf_target_path) pairs using docx2pdf."""
    try:
        from docx2pdf import convert as docx2pdf_convert
    except ImportError:
        return
    tmp_dir = tempfile.mkdtemp(prefix="exam_pdf_")
    mapping = []
    try:
        for idx, (dx, px) in enumerate(pairs, 1):
            if not os.path.exists(dx):
                continue
            tmp_dx = os.path.join(tmp_dir, f"file{idx}.docx")
            tmp_px = os.path.join(tmp_dir, f"file{idx}.pdf")
            shutil.copy2(dx, tmp_dx)
            mapping.append((tmp_px, px))
        docx2pdf_convert(tmp_dir)
        for tmp_px, target_px in mapping:
            if os.path.exists(tmp_px):
                os.makedirs(os.path.dirname(target_px), exist_ok=True)
                shutil.copy2(tmp_px, target_px)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 5. Draft (auto-save) helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def save_draft(exam_text, ans_text, title):
    try:
        with open(DRAFT_FILE, "w", encoding="utf-8") as f:
            json.dump({"exam": exam_text, "answer": ans_text, "title": title,
                        "saved_at": datetime.datetime.now().isoformat()}, f, ensure_ascii=False)
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
# 6. History helpers
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
# 7. Question-bank (JSON flat-file) helpers
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def save_to_qbank(questions, tag="", answers_map=None):
    """Append parsed questions to the JSON question bank.
    answers_map: {original_num: answer_text} dict (optional)
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
    if not os.path.exists(QBANK_FILE):
        return []
    try:
        with open(QBANK_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def clear_qbank():
    if os.path.exists(QBANK_FILE):
        os.remove(QBANK_FILE)
