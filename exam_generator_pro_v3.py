"""
사탐/경제 모의고사 자동 생성기 PRO v4.0 (Advanced Architecture)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
v3 대비 개선:
  - Drag & Drop 지원 (.txt, .docx 텍스트 자동 추출)
  - 사용자 로컬 이미지 삽입 ([IMAGE:경로] 태그 지원)
  - 확장 경제 그래프 (조세 부과, 관세, 외부효과, 무차별곡선) 추가 및 영문/기호 기반 렌더링
  - 문제 은행 다중 필터 검색 (검색어 조건 및 정답 유무)
  - 문항 메타데이터(난이도, 오답률) 편집 및 DB 동기화 시스템
v2 대비 (기존 유지):
  - 클래스 기반 아키텍처 / 실시간 미리보기 / 배치 생성
  - 문제 섞기 / 확장 그래프 (PPF, AS-AD, 로렌츠, 필립스)
  - 문제 은행 + OMR 카드 + 히스토리 + 자동 임시저장
  - AI 명령어 탭 + 정답/해설 자동 매칭
"""
import os
import re
import hashlib
import shutil
import tempfile
import json
import datetime
import random
import tkinter as tk
from tkinter import messagebox, filedialog, ttk, simpledialog
from collections import OrderedDict
import textwrap
import docx
from docx.shared import Cm, Pt, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import numpy as np
import platform
try:
    from docx2pdf import convert as docx2pdf_convert
except ImportError:
    docx2pdf_convert = None
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError:
    TkinterDnD = None

import exam_db

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 0. AI MASTER PROMPT
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
# 1. 환경 설정
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


def _setup_korean_font():
    try:
        plt.style.use("seaborn-v0_8-whitegrid")
    except Exception:
        pass
    names = {f.name for f in fm.fontManager.ttflist}
    system = platform.system()
    if system == "Windows":
        for c in ("Malgun Gothic", "NanumGothic", "Gulim"):
            if c in names:
                plt.rcParams["font.family"] = c
                break
    elif system == "Darwin":
        plt.rcParams["font.family"] = "AppleGothic"
    else:
        plt.rcParams["font.family"] = "NanumGothic"
    plt.rcParams["axes.unicode_minus"] = False

_setup_korean_font()


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
    if docx2pdf_convert is None:
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
# 2. 확장 그래프 엔진 (PPF, AS-AD, 로렌츠 추가)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _draw_base_axes(ax, xlabel="Q", ylabel="P"):
    ax.grid(True, linestyle="--", color="lightgray", alpha=0.7, zorder=0)
    ax.spines["right"].set_visible(False)
    ax.spines["top"].set_visible(False)
    ax.set_xlabel(xlabel, loc="right", fontsize=11, fontweight="bold", fontstyle="italic")
    ax.set_ylabel(ylabel, loc="top", fontsize=11, fontweight="bold", fontstyle="italic", rotation=0)


def _draw_sd_subplot(ax, shift_match):
    _draw_base_axes(ax)
    q = [0, 10]
    ax.plot(q, [0, 10], "k-", lw=2, zorder=3)
    ax.plot(q, [10, 0], "k-", lw=2, zorder=3)
    ax.text(9.5, 9.5, "S", fontsize=12, fontweight="bold")
    ax.text(9.5, 0.5, "D", fontsize=12, fontweight="bold")
    ax.set_xticks([])
    ax.set_yticks([])
    if not shift_match:
        return
    curve = shift_match.group(1).upper()
    direction = shift_match.group(2).upper()
    color_map = {"S": "darkred", "D": "blue"}
    color = color_map.get(curve, "gray")
    offset = 2 if direction == "RIGHT" else -2
    arrow_dx = 2 if direction == "RIGHT" else -2
    if curve == "S":
        new_line = ([offset, 10 + offset], [0, 10])
    else:
        new_line = ([offset, 10 + offset], [10, 0])
    ax.plot(new_line[0], new_line[1], color=color, linestyle="--", lw=2, zorder=3)
    label_x = new_line[0][-1] - 0.5
    label_y = new_line[1][-1] + (0.5 if curve == "S" else -0.5)
    ax.text(label_x, label_y, f"{curve}'", fontsize=12, fontweight="bold", color=color)
    ax.annotate("", xy=(5 + arrow_dx / 2, 5), xytext=(5 - arrow_dx / 2, 5),
                arrowprops=dict(arrowstyle="->", color="gray", lw=1.5))


def _draw_ppf(ax, tag_str):
    _draw_base_axes(ax, xlabel="재화 X", ylabel="재화 Y")
    ax.set_xticks([])
    ax.set_yticks([])
    t = np.linspace(0, 1, 100)
    x = 10 * (1 - t ** 2) ** 0.5
    y = 10 * t
    ax.plot(x, y, "b-", lw=2.5, zorder=3, label="PPF")
    ax.fill_between(x, y, alpha=0.05, color="blue")
    ax.text(x[0] + 0.3, y[0] + 0.3, "A", fontsize=11, fontweight="bold")
    ax.text(x[-1] + 0.3, y[-1] - 0.5, "B", fontsize=11, fontweight="bold")
    shift = re.search(r"PPF_SHIFT\s*=\s*(OUT|IN)", tag_str, re.IGNORECASE)
    if shift:
        factor = 1.3 if shift.group(1).upper() == "OUT" else 0.7
        x2 = factor * 10 * (1 - t ** 2) ** 0.5
        y2 = factor * 10 * t
        ax.plot(x2, y2, "r--", lw=2, zorder=3, label=f"PPF' ({shift.group(1)})")
        mid = len(t) // 2
        ax.annotate("", xy=(x2[mid], y2[mid]), xytext=(x[mid], y[mid]),
                    arrowprops=dict(arrowstyle="->", color="gray", lw=1.5))
    ax.legend(fontsize=9)


def _draw_asad(ax, tag_str):
    _draw_base_axes(ax, xlabel="Y (실질GDP)", ylabel="P (물가)")
    ax.set_xticks([])
    ax.set_yticks([])
    q = [0, 10]
    ax.plot(q, [10, 0], "b-", lw=2, zorder=3)
    ax.text(9.5, 0.5, "AD", fontsize=11, fontweight="bold", color="blue")
    ax.plot(q, [2, 8], "r-", lw=2, zorder=3)
    ax.text(9.5, 8.3, "SRAS", fontsize=10, fontweight="bold", color="red")
    ax.axvline(x=6, color="green", lw=2, linestyle="-", zorder=3)
    ax.text(6.2, 9.5, "LRAS", fontsize=10, fontweight="bold", color="green")
    ad_shift = re.search(r"Shift_AD\s*=\s*(Right|Left)", tag_str, re.IGNORECASE)
    if ad_shift:
        d = ad_shift.group(1).upper()
        offset = 2 if d == "RIGHT" else -2
        ax.plot([offset, 10 + offset], [10, 0], "b--", lw=2, zorder=3)
        ax.text(9.5 + offset, 0.5, "AD'", fontsize=11, fontweight="bold", color="darkblue")
    sras_shift = re.search(r"Shift_SRAS\s*=\s*(Right|Left)", tag_str, re.IGNORECASE)
    if sras_shift:
        d = sras_shift.group(1).upper()
        offset = 2 if d == "RIGHT" else -2
        ax.plot([offset, 10 + offset], [2, 8], "r--", lw=2, zorder=3)
        ax.text(9.5 + offset, 8.3, "SRAS'", fontsize=10, fontweight="bold", color="darkred")


def _draw_lorenz(ax, tag_str):
    _draw_base_axes(ax, xlabel="인구 누적 비율(%)", ylabel="소득 누적 비율(%)")
    x = np.linspace(0, 1, 100)
    ax.plot(x, x, "k--", lw=1.5, label="완전평등선", zorder=3)
    gini = re.search(r"GINI\s*=\s*([0-9.]+)", tag_str, re.IGNORECASE)
    exp = 2.0
    if gini:
        g = float(gini.group(1))
        exp = max(1.1, 1 / (1 - g + 0.01))
    y = x ** exp
    ax.plot(x, y, "b-", lw=2.5, label="로렌츠 곡선", zorder=3)
    ax.fill_between(x, x, y, alpha=0.15, color="blue")
    ax.text(0.6, 0.3, "불평등\n면적", fontsize=9, ha="center", color="blue")
    compare = re.search(r"COMPARE_GINI\s*=\s*([0-9.]+)", tag_str, re.IGNORECASE)
    if compare:
        g2 = float(compare.group(1))
        exp2 = max(1.1, 1 / (1 - g2 + 0.01))
        y2 = x ** exp2
        ax.plot(x, y2, "r--", lw=2, label=f"비교 (Gini={g2})", zorder=3)
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.legend(fontsize=8, loc="upper left")


def _draw_phillips(ax, tag_str):
    _draw_base_axes(ax, xlabel="실업률(%)", ylabel="인플레이션(%)")
    ax.set_xticks([])
    ax.set_yticks([])
    x = np.linspace(1, 10, 100)
    y = 8 / x + 0.5
    ax.plot(x, y, "b-", lw=2.5, zorder=3, label="단기 필립스")
    shift = re.search(r"PHILLIPS_SHIFT\s*=\s*(UP|DOWN)", tag_str, re.IGNORECASE)
    if shift:
        d = shift.group(1).upper()
        offset = 2 if d == "UP" else -2
        y2 = 8 / x + 0.5 + offset
        ax.plot(x, y2, "r--", lw=2, zorder=3, label=f"이동 ({d})")
    ax.legend(fontsize=9)


def _draw_tax(ax, tag_str):
    """조세 부과 그래프 (Tax incidence) — 영문 라벨."""
    _draw_base_axes(ax, xlabel="Q", ylabel="P")
    q = np.linspace(0, 10, 100)
    d = 10 - q
    s = q
    tax = 2
    s_tax = q + tax
    ax.plot(q, d, "b-", lw=2, label="D", zorder=3)
    ax.plot(q, s, "r-", lw=2, label="S", zorder=3)
    ax.plot(q, s_tax, "r--", lw=2, label="S + Tax", zorder=3)
    # 균형점
    q_eq1, p_eq1 = 5, 5
    q_eq2, p_eq2, p_seller = 4, 6, 4
    ax.plot(q_eq1, p_eq1, "ko", ms=5, zorder=4)
    ax.plot(q_eq2, p_eq2, "ko", ms=5, zorder=4)
    ax.plot([0, q_eq2, q_eq2], [p_eq2, p_eq2, 0], "k:", lw=1, zorder=2)
    ax.plot([0, q_eq2], [p_seller, p_seller], "k:", lw=1, zorder=2)
    ax.fill_between([0, q_eq2], p_seller, p_eq2, alpha=0.1, color="orange")
    ax.text(q_eq2 + 0.15, p_eq2 + 0.3, "E'", fontsize=10, fontweight="bold")
    ax.text(q_eq1 + 0.15, p_eq1 + 0.3, "E", fontsize=10, fontweight="bold")
    ax.text(0.3, (p_eq2 + p_seller) / 2, "Tax\nRevenue", fontsize=8, ha="left", va="center", color="orange")
    ax.legend(fontsize=9, loc="upper right")


def _draw_tariff(ax, tag_str):
    """관세 그래프 (Tariff) — 영문 라벨."""
    _draw_base_axes(ax, xlabel="Q", ylabel="P")
    q = np.linspace(0, 10, 100)
    d = 10 - q
    s = q
    p_w = 3
    tariff = 2
    p_t = p_w + tariff
    ax.plot(q, d, "b-", lw=2, label="D (Domestic)", zorder=3)
    ax.plot(q, s, "r-", lw=2, label="S (Domestic)", zorder=3)
    ax.axhline(p_w, color="green", linestyle="-", lw=1.5, label="Pw (World Price)", zorder=3)
    ax.axhline(p_t, color="purple", linestyle="--", lw=1.5, label="Pw + Tariff", zorder=3)
    # 수량 표시
    qd_w = 10 - p_w
    qs_w = p_w
    qd_t = 10 - p_t
    qs_t = p_t
    ax.fill_between([qs_w, qs_t], p_w, p_t, alpha=0.08, color="red")
    ax.fill_between([qd_t, qd_w], p_w, p_t, alpha=0.08, color="red")
    ax.fill_between([qs_t, qd_t], p_w, p_t, alpha=0.15, color="purple")
    ax.text((qs_t + qd_t) / 2, (p_w + p_t) / 2, "Tariff\nRevenue", fontsize=8,
            ha="center", va="center", color="purple")
    ax.legend(fontsize=8, loc="upper right")


def _draw_externality(ax, tag_str):
    """외부효과 그래프 (Externality — negative production) — 영문 라벨."""
    _draw_base_axes(ax, xlabel="Q", ylabel="P / Cost / Benefit")
    q = np.linspace(0, 10, 100)
    pmb = 10 - q          # PMB = SMB (부정적 외부효과: 생산 측)
    pmc = q + 1            # Private Marginal Cost
    smc = q + 4            # Social Marginal Cost (외부비용 포함)
    ax.plot(q, pmb, "b-", lw=2, label="PMB = SMB", zorder=3)
    ax.plot(q, pmc, "r-", lw=2, label="PMC", zorder=3)
    ax.plot(q, smc, "r--", lw=2, label="SMC", zorder=3)
    # 사회적 최적 vs 시장 균형
    q_market = 4.5  # PMB = PMC 교차
    q_social = 3.0  # PMB = SMC 교차
    ax.fill_between(q, pmc, smc, where=(q >= q_social) & (q <= q_market),
                    color="gray", alpha=0.3, zorder=2)
    ax.text((q_social + q_market) / 2, (pmc[50] + smc[50]) / 2 - 0.3,
            "DWL", fontsize=10, fontweight="bold", ha="center")
    ax.plot([q_market, q_market], [0, 10 - q_market], "k:", lw=1)
    ax.plot([q_social, q_social], [0, 10 - q_social], "k:", lw=1)
    ax.text(q_market, -0.5, "Qm", fontsize=9, ha="center")
    ax.text(q_social, -0.5, "Qs", fontsize=9, ha="center")
    ax.legend(fontsize=8, loc="upper right")


def _draw_indifference(ax, tag_str):
    """무차별곡선 + 예산선 그래프 (Indifference Curves) — 영문 라벨."""
    _draw_base_axes(ax, xlabel="Good X", ylabel="Good Y")
    ax.set_xticks([])
    ax.set_yticks([])
    x = np.linspace(0.5, 10, 100)
    u1 = 12 / x
    u2 = 20 / x
    budget = 10 - x
    ax.plot(x, u1, "b-", lw=2, label="U1", zorder=3)
    ax.plot(x, u2, "b--", lw=2, label="U2", zorder=3)
    ax.plot(x[x <= 10], np.maximum(budget[x <= 10], 0), "k-", lw=1.5, label="Budget Line", zorder=3)
    # 최적점 (U1과 예산선의 접점 근사)
    # U1 = 12/x, Budget = 10-x → 12/x = 10-x → x^2 - 10x + 12 = 0
    x_opt = (10 - np.sqrt(100 - 48)) / 2  # ≈ 1.35... 더 현실적인 근사
    x_opt2 = (10 + np.sqrt(100 - 48)) / 2  # ≈ 8.65... 이게 맞음 (내부해)
    # 실제로 근사치는 약 x=3.5, y=3.4가 합리적
    ax.plot(3.46, 3.46, "ro", ms=6, zorder=5)
    ax.text(3.8, 3.6, "Optimal", fontsize=9, fontweight="bold", color="red")
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 11)
    ax.legend(fontsize=9, loc="upper right")


def process_and_draw_graph(tag_str, save_path):
    clean = re.sub(r"[\$\\\[\]\|]", " ", tag_str)
    upper = clean.upper()
    # v4 확장 그래프 (영문/기호 라벨)
    if "TAX" in upper and "TARIFF" not in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_tax(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    if "TARIFF" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_tariff(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    if "EXTERNALITY" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_externality(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    if "INDIFFERENCE" in upper or ("CURVE" in upper and "PPF" not in upper
                                    and "LORENZ" not in upper and "PHILLIPS" not in upper):
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_indifference(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    # 기존 v3 그래프
    if "PPF" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_ppf(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    if "ASAD" in upper or "AS_AD" in upper or "AS-AD" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_asad(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    if "LORENZ" in upper or "GINI" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_lorenz(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    if "PHILLIPS" in upper:
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_phillips(ax, clean)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    if "DUAL" in upper or "LEFT:" in upper:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(7, 3.5))
        left_m = re.search(r"Left:\s*Shift_([SD])\s*=\s*(Right|Left)", clean, re.IGNORECASE)
        right_m = re.search(r"Right:\s*Shift_([SD])\s*=\s*(Right|Left)", clean, re.IGNORECASE)
        _draw_sd_subplot(ax1, left_m)
        _draw_sd_subplot(ax2, right_m)
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    qd_m = re.search(r"Qd\s*=\s*([0-9, ]+)", clean, re.IGNORECASE)
    qs_m = re.search(r"Qs\s*=\s*([0-9, ]+)", clean, re.IGNORECASE)
    p_m = re.search(r"P\s*=\s*([0-9, ]+)", clean, re.IGNORECASE)
    shift_s_m = re.search(r"Shift_S\s*=\s*([-0-9]+)", clean, re.IGNORECASE)
    shift_d_m = re.search(r"Shift_D\s*=\s*([-0-9]+)", clean, re.IGNORECASE)
    ceiling_m = re.search(r"Price_Ceiling\s*=\s*([0-9]+)", clean, re.IGNORECASE)
    floor_m = re.search(r"Price_Floor\s*=\s*([0-9]+)", clean, re.IGNORECASE)
    if qd_m and qs_m:
        qd = list(map(int, qd_m.group(1).split(",")))
        qs = list(map(int, qs_m.group(1).split(",")))
        p_vals = list(map(int, p_m.group(1).split(","))) if p_m else [10 * (i + 1) for i in range(len(qd))]
        fig, ax = plt.subplots(figsize=(5, 4))
        _draw_base_axes(ax)
        ax.plot(qd, p_vals, color="blue", lw=2, zorder=3)
        ax.plot(qs, p_vals, color="red", lw=2, zorder=3)
        ax.text(qd[-1], p_vals[-1] - max(p_vals) * 0.05, "D", fontsize=12, color="blue", fontweight="bold")
        ax.text(qs[-1], p_vals[-1] + max(p_vals) * 0.05, "S", fontsize=12, color="red", fontweight="bold")
        if shift_s_m:
            s = int(shift_s_m.group(1))
            sp = [p + s for p in p_vals]
            ax.plot(qs, sp, color="darkred", linestyle="--", lw=2, zorder=3)
            ax.text(qs[-1], sp[-1] + max(p_vals) * 0.05, "S'", fontsize=12, color="darkred", fontweight="bold")
        if shift_d_m:
            s = int(shift_d_m.group(1))
            sp = [p + s for p in p_vals]
            ax.plot(qd, sp, color="darkblue", linestyle="--", lw=2, zorder=3)
            ax.text(qd[-1], sp[-1] + max(p_vals) * 0.05, "D'", fontsize=12, color="darkblue", fontweight="bold")
        if ceiling_m:
            c = int(ceiling_m.group(1))
            ax.axhline(y=c, color="purple", linestyle="-.", lw=2, zorder=4)
            ax.text(min(qd), c + max(p_vals) * 0.03, "Price Ceiling", color="purple", fontweight="bold", fontstyle="italic")
        if floor_m:
            f = int(floor_m.group(1))
            ax.axhline(y=f, color="green", linestyle="-.", lw=2, zorder=4)
            ax.text(min(qd), f + max(p_vals) * 0.03, "Price Floor", color="green", fontweight="bold", fontstyle="italic")
        plt.tight_layout()
        fig.savefig(save_path, dpi=300)
        plt.close(fig)
        return save_path
    return None

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 3. 텍스트 파싱 엔진
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CIRCLE_NUMS = "①②③④⑤"


def parse_exam_text(raw_text):
    """원시 텍스트 -> 구조화된 문제 리스트."""
    raw_text = raw_text.lstrip("\ufeff")
    raw_text = raw_text.replace("\r\n", "\n").replace("\r", "\n")
    raw_text = raw_text.replace("\u00a0", " ")
    raw_text = raw_text.replace("\u3000", " ")
    raw_text = re.sub(r"```[a-zA-Z]*\n?", "", raw_text)
    raw_text = raw_text.replace("```", "").replace("**", "").replace("$", "")

    SPLIT_PAT = r"\n(?=(?:Q|문|문항)?\s{0,3}\d{1,3}[\.\):\s]\s*[가-힣A-Za-z<\(\[（「【])"
    MATCH_PAT = r"^\s*(?:Q|문|문항)?\s{0,3}([0-9]{1,3})[\.\):]?\s+(.*)"

    splits = re.split(SPLIT_PAT, "\n" + raw_text)
    result = []

    for chunk in splits:
        if not chunk.strip():
            continue
        m = re.match(MATCH_PAT, chunk, re.DOTALL | re.IGNORECASE)
        if not m:
            continue
        q_num = m.group(1)
        body = m.group(2)

        # 이미지 태그 추출 [IMAGE:경로]
        image_path = None
        img_match = re.search(r"\[IMAGE\s*:\s*(.+?)\]", body, re.IGNORECASE)
        if img_match:
            image_path = img_match.group(1).strip()
            body = body.replace(img_match.group(0), "").strip()

        # 선택지 추출
        choices = []
        c_start = re.search(r"(①|1\s*\))", body)
        if c_start:
            choices_str = body[c_start.start():]
            body = body[:c_start.start()].strip()
            for i in range(1, 6):
                circle = CIRCLE_NUMS[i - 1]
                if i < 5:
                    next_circle = CIRCLE_NUMS[i]
                    pat = f"({re.escape(circle)}|{i}\\s*\\)).*?(?={re.escape(next_circle)}|{i+1}\\s*\\)|$)"
                else:
                    pat = f"({re.escape(circle)}|{i}\\s*\\)).*"
                cm = re.search(pat, choices_str, re.DOTALL)
                if cm:
                    val = re.sub(f"^({re.escape(circle)}|{i}\\s*\\))\\s*", "", cm.group(0)).replace("\n", " ").strip()
                    choices.append(val)
                else:
                    choices.append("")

        # 그래프 태그 추출
        lines = body.split("\n")
        clean_lines = []
        graph_parts = []
        in_graph = False
        for line in lines:
            up = line.upper()
            if re.search(r"\[\s*(?:DUAL_)?(?:GRAPH|PPF|ASAD|AS.AD|LORENZ|PHILLIPS|TAX|TARIFF|EXTERNALITY|INDIFFERENCE)", up):
                in_graph = True
                graph_parts.append(line)
                if line.count("[") == line.count("]") and "]" in line:
                    in_graph = False
            elif in_graph:
                graph_parts.append(line)
                if "]" in line:
                    in_graph = False
            elif re.search(r"^\s*[\|\[\$]?\s*(?:Qd|Qs|P|Shift_[SD]|Price_|PPF_SHIFT|GINI|COMPARE_GINI|PHILLIPS_SHIFT|Shift_AD|Shift_SRAS)\s*=", line, re.IGNORECASE):
                graph_parts.append(line)
            else:
                clean_lines.append(line)
        graph_tag = " ".join(graph_parts) if graph_parts else None
        body = "\n".join(clean_lines).strip()

        # 제시문 추출
        jesi = ""
        jm = re.search(r"<제시문>\s*(.*?)\s*</제시문>", body, re.DOTALL | re.IGNORECASE)
        if jm:
            jesi = jm.group(1).strip()
            body = body.replace(jm.group(0), "").strip()

        # 보기 추출
        bogi = ""
        bogi_tag = re.search(r"^\s*[\[<]\s*보\s*기\s*[\]>][ \t]*\n?", body, re.MULTILINE)
        if bogi_tag:
            bogi = body[bogi_tag.end():].strip()
            body = body[:bogi_tag.start()].strip()
        else:
            bogi_lines = re.findall(r"^[ㄱㄴㄷㄹㅁㅂ]\s*[\.\)].*", body, re.MULTILINE)
            if bogi_lines:
                bogi = "\n".join(bogi_lines)
                for bl in bogi_lines:
                    body = body.replace(bl, "")
                body = body.strip()

        body = re.sub(r"[\[<]\s*보\s*기\s*[\]>]", "보기", body).strip()

        # (가)(나)(다) 블록 -> 제시문
        if not jesi:
            gana_pat = r"^\s*[\(（]([가나다라마바사아자차카타파하])\s*[\)）]\s+.+"
            gana_lines = re.findall(gana_pat, body, re.MULTILINE)
            if len(gana_lines) >= 2:
                gana_full = re.findall(
                    r"^\s*[\(（][가나다라마바사아자차카타파하]\s*[\)）]\s+.+",
                    body, re.MULTILINE
                )
                jesi = "\n".join(line.strip() for line in gana_full)
                for gl in gana_full:
                    body = body.replace(gl, "")
                body = body.strip()

        # 표 추출
        tables = []
        final_lines = []
        current_table = []
        for line in body.split("\n"):
            ls = line.strip()
            if "|" in ls and not re.match(r"^[-:\s\|\+]+$", ls):
                current_table.append([c.strip() for c in ls.strip("|").split("|")])
            elif "|" in ls and re.match(r"^[-:\s\|\+]+$", ls):
                continue
            else:
                if current_table:
                    tables.append(current_table)
                    final_lines.append("[TABLE_PLACEHOLDER]")
                    current_table = []
                final_lines.append(line)
        if current_table:
            tables.append(current_table)
            final_lines.append("[TABLE_PLACEHOLDER]")
        body = re.sub(r"\n{3,}", "\n\n", "\n".join(final_lines).strip())

        result.append({
            "num": q_num,
            "text": body,
            "tables": tables,
            "bogi": bogi,
            "jesi": jesi,
            "graph_tag": graph_tag,
            "image_path": image_path,
            "choices": choices,
        })

    return result


def parse_answer_by_question(raw_ans):
    """정답/해설 텍스트를 문항번호별로 분리하여 딕셔너리로 반환.
    반환: {"1": "1번 정답: ③\\n[해설] ...", "2": "2번 정답: ⑤\\n...", ...}
    """
    if not raw_ans or not raw_ans.strip():
        return {}

    result = {}
    # 문항 번호 기준으로 분리
    # 패턴: 줄 시작 + 숫자 + (번|.|)| 등
    split_pat = r"\n(?=\s*\d{1,3}\s*(?:번|[\.\)])\s*)"
    chunks = re.split(split_pat, "\n" + raw_ans)

    for chunk in chunks:
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.match(r"^\s*(\d{1,3})\s*(?:번|[\.\)])\s*(.*)", chunk, re.DOTALL)
        if m:
            q_num = m.group(1).lstrip("0") or "0"
            result[q_num] = chunk
    return result

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 4. Word 문서 생성 엔진
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _apply_doc_style(doc, font_name, font_size):
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(int(font_size))
    style._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(1)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    sec = doc.sections[0]
    sec.top_margin = sec.bottom_margin = Cm(1.2)
    sec.left_margin = sec.right_margin = Cm(1.3)


def _set_two_columns(section):
    try:
        spr = section._sectPr
        existing = spr.xpath("./w:cols")
        elem = existing[0] if existing else OxmlElement("w:cols")
        elem.set(qn("w:num"), "2")
        elem.set(qn("w:space"), "708")
        elem.set(qn("w:sep"), "1")
        if not existing:
            spr.append(elem)
    except Exception:
        pass


def _set_cell_bg(cell, color):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), color)
    tcPr.append(shd)


def _add_watermark(doc, text):
    if not text:
        return
    for section in doc.sections:
        header = section.header
        header.is_linked_to_previous = False
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.font.size = Pt(48)
        run.font.color.rgb = RGBColor(220, 220, 220)


def _add_header_footer(doc, header_text="", footer_text=""):
    for section in doc.sections:
        if header_text:
            h = section.header
            h.is_linked_to_previous = False
            hp = h.paragraphs[0] if h.paragraphs else h.add_paragraph()
            hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            hr = hp.add_run(header_text)
            hr.font.size = Pt(8)
            hr.font.color.rgb = RGBColor(150, 150, 150)
        if footer_text:
            f = section.footer
            f.is_linked_to_previous = False
            fp = f.paragraphs[0] if f.paragraphs else f.add_paragraph()
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            fr = fp.add_run(footer_text)
            fr.font.size = Pt(8)
            fr.font.color.rgb = RGBColor(150, 150, 150)


def _add_page_number_field(paragraph, font_size_pt=8, font_color_hex="969696"):
    """w:fldSimple을 사용해 실제 페이지 번호 필드를 삽입."""
    prefix_run = paragraph.add_run("Page ")
    prefix_run.font.size = Pt(font_size_pt)
    prefix_run.font.color.rgb = RGBColor(
        int(font_color_hex[0:2], 16),
        int(font_color_hex[2:4], 16),
        int(font_color_hex[4:6], 16),
    )

    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " PAGE \\* MERGEFORMAT ")
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), str(font_size_pt * 2))
    rPr.append(sz)
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(font_size_pt * 2))
    rPr.append(szCs)
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), font_color_hex)
    rPr.append(color_el)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.text = "1"
    r.append(t)
    fld.append(r)
    paragraph._element.append(fld)


def _add_header_footer_v2(doc, hf_config, logo_path=""):
    """향상된 머리글/바닥글 — 커스텀 필드 + OxmlElement 기반 페이지 번호.
    모든 section(continuous 포함)에 적용."""
    academy = hf_config.get("academy_name", "")
    exam_name = hf_config.get("exam_name", "")
    contact = hf_config.get("contact_info", "")
    copyright_text = hf_config.get("copyright_text", "")
    show_page = hf_config.get("show_page_number", False)
    show_date = hf_config.get("show_date", False)

    date_str = datetime.datetime.now().strftime("%Y-%m-%d") if show_date else ""
    header_parts = [p for p in [academy, exam_name, date_str] if p]
    header_text = " | ".join(header_parts)
    footer_parts = [p for p in [copyright_text, contact] if p]
    footer_text = " | ".join(footer_parts)

    has_header = bool(header_text) or show_page
    has_footer = bool(footer_text) or show_page

    for section in doc.sections:
        if has_header:
            h = section.header
            h.is_linked_to_previous = False
            hp = h.paragraphs[0] if h.paragraphs else h.add_paragraph()
            hp.clear()
            hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if header_text:
                hr = hp.add_run(header_text)
                hr.font.size = Pt(8)
                hr.font.color.rgb = RGBColor(150, 150, 150)
            if show_page:
                if header_text:
                    sep = hp.add_run(" | ")
                    sep.font.size = Pt(8)
                    sep.font.color.rgb = RGBColor(150, 150, 150)
                _add_page_number_field(hp, 8, "969696")

        if has_footer:
            f = section.footer
            f.is_linked_to_previous = False
            fp = f.paragraphs[0] if f.paragraphs else f.add_paragraph()
            fp.clear()
            fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if footer_text:
                fr = fp.add_run(footer_text)
                fr.font.size = Pt(8)
                fr.font.color.rgb = RGBColor(150, 150, 150)
            if show_page:
                if footer_text:
                    sep = fp.add_run(" | ")
                    sep.font.size = Pt(8)
                    sep.font.color.rgb = RGBColor(150, 150, 150)
                _add_page_number_field(fp, 8, "969696")


def _tight_para(doc, left_indent=None, space_before=0, space_after=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    if left_indent:
        pf.left_indent = left_indent
    return p


def _tight_table_cell(cell, font_size=None):
    for p in cell.paragraphs:
        pf = p.paragraph_format
        pf.space_before = Pt(1)
        pf.space_after = Pt(1)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        if font_size:
            for run in p.runs:
                run.font.size = Pt(font_size)


def create_exam_docx(target_dir, filename, exam_data, font_name, font_size,
                     logo_path, title, watermark="", header="", footer="",
                     numbering_start=1, hf_config=None):
    fs = int(font_size)
    doc = docx.Document()
    _apply_doc_style(doc, font_name, font_size)
    if watermark:
        _add_watermark(doc, watermark)
    if hf_config and any(hf_config.get(k) for k in ("academy_name", "exam_name", "show_page_number", "show_date", "copyright_text", "contact_info")):
        _add_header_footer_v2(doc, hf_config, logo_path=logo_path)
    elif header or footer:
        _add_header_footer(doc, header, footer)

    if logo_path and os.path.exists(logo_path):
        lp = _tight_para(doc, space_after=2)
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lp.add_run().add_picture(logo_path, height=Cm(2.5))
    tp = doc.add_heading(title, 1)
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_after = Pt(4)

    sep = _tight_para(doc, space_before=0, space_after=6)
    sep_run = sep.add_run("━" * 60)
    sep_run.font.size = Pt(6)
    sep_run.font.color.rgb = RGBColor(180, 180, 180)
    sep.alignment = WD_ALIGN_PARAGRAPH.CENTER

    new_sec = doc.add_section(WD_SECTION.CONTINUOUS)
    _set_two_columns(new_sec)

    for q_idx, q in enumerate(exam_data):
        display_num = numbering_start + q_idx

        parts = q["text"].split("[TABLE_PLACEHOLDER]")
        for idx, part in enumerate(parts):
            if part.strip() or idx == 0:
                p = _tight_para(doc, space_before=8 if idx == 0 else 1, space_after=2)
                if idx == 0:
                    num_run = p.add_run(f"{display_num}. ")
                    num_run.bold = True
                    num_run.font.size = Pt(fs + 1)
                text = part.strip()
                if text:
                    body_run = p.add_run(text)
                    body_run.font.size = Pt(fs)
            if idx < len(q["tables"]):
                td = q["tables"][idx]
                if td and td[0]:
                    tbl = doc.add_table(rows=len(td), cols=len(td[0]))
                    tbl.style = "Table Grid"
                    for ri, row in enumerate(td):
                        for ci, cell_text in enumerate(row):
                            if ci < len(tbl.columns):
                                cell = tbl.cell(ri, ci)
                                cell.text = cell_text
                                _tight_table_cell(cell, font_size=fs - 1)
                                if ri == 0:
                                    _set_cell_bg(cell, "E7E6E6")
                                    if cell.paragraphs[0].runs:
                                        cell.paragraphs[0].runs[0].bold = True
                                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        if q["jesi"]:
            jt = doc.add_table(rows=1, cols=1)
            jt.style = "Table Grid"
            cell = jt.cell(0, 0)
            _set_cell_bg(cell, "F9F9F9")
            for ji, jline in enumerate(q["jesi"].split("\n")):
                if ji == 0:
                    cell.paragraphs[0].text = jline.strip()
                else:
                    cell.add_paragraph(jline.strip())
            _tight_table_cell(cell, font_size=fs - 1)

        # 사용자 로컬 이미지 삽입 ([IMAGE:경로])
        if q.get("image_path") and os.path.exists(q["image_path"]):
            try:
                img_p = _tight_para(doc, space_before=2, space_after=2)
                img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                img_p.add_run().add_picture(q["image_path"], width=Cm(12.0))
            except Exception:
                pass

        if q.get("graph_tag"):
            tmp_img = os.path.join(target_dir, f"_tmp_graph_{q['num']}.png")
            img = process_and_draw_graph(q["graph_tag"], tmp_img)
            if img and os.path.exists(img):
                ip = _tight_para(doc, space_before=2, space_after=2)
                ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
                ip.add_run().add_picture(img, width=Cm(6.5))
                os.remove(img)

        if q["bogi"]:
            bp = _tight_para(doc, space_before=3, space_after=1)
            bp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            br = bp.add_run("< 보 기 >")
            br.bold = True
            br.font.size = Pt(fs)
            bt = doc.add_table(rows=1, cols=1)
            bt.style = "Table Grid"
            cell = bt.cell(0, 0)
            _set_cell_bg(cell, "FAFAFA")
            for bi, bline in enumerate(q["bogi"].split("\n")):
                if bi == 0:
                    cell.paragraphs[0].text = bline.strip()
                else:
                    cell.add_paragraph(bline.strip())
            _tight_table_cell(cell, font_size=fs - 1)

        if q["choices"]:
            valid = [(i, c) for i, c in enumerate(q["choices"]) if c]
            if valid:
                max_len = max(len(c) for _, c in valid)
                if max_len <= 8 and len(valid) == 5:
                    cp = _tight_para(doc, left_indent=Cm(0.3), space_before=3, space_after=1)
                    parts_str = "   ".join(f"{CIRCLE_NUMS[i]} {c}" for i, c in valid)
                    cr = cp.add_run(parts_str)
                    cr.font.size = Pt(fs)
                elif max_len <= 18 and len(valid) >= 4:
                    first_row = valid[:3]
                    second_row = valid[3:]
                    cp1 = _tight_para(doc, left_indent=Cm(0.3), space_before=3, space_after=0)
                    r1 = cp1.add_run("   ".join(f"{CIRCLE_NUMS[i]} {c}" for i, c in first_row))
                    r1.font.size = Pt(fs)
                    if second_row:
                        cp2 = _tight_para(doc, left_indent=Cm(0.3), space_before=0, space_after=1)
                        r2 = cp2.add_run("   ".join(f"{CIRCLE_NUMS[i]} {c}" for i, c in second_row))
                        r2.font.size = Pt(fs)
                else:
                    for ci, (i, c) in enumerate(valid):
                        cp = _tight_para(doc, left_indent=Cm(0.5),
                                         space_before=1 if ci == 0 else 0, space_after=0)
                        cr = cp.add_run(f"{CIRCLE_NUMS[i]} {c}")
                        cr.font.size = Pt(fs)

        if q_idx < len(exam_data) - 1:
            div = _tight_para(doc, space_before=4, space_after=2)
            div_run = div.add_run("─" * 35)
            div_run.font.size = Pt(4)
            div_run.font.color.rgb = RGBColor(210, 210, 210)

    dx, px = build_paths(target_dir, filename)
    doc.save(dx)
    return dx, px


def create_answer_docx(target_dir, filename, raw_ans, exam_data, font_name, font_size,
                       logo_path, title, numbering_start=1, hf_config=None):
    doc = docx.Document()
    _apply_doc_style(doc, font_name, font_size)
    if hf_config and any(hf_config.get(k) for k in ("academy_name", "exam_name", "show_page_number", "show_date", "copyright_text", "contact_info")):
        _add_header_footer_v2(doc, hf_config, logo_path=logo_path)

    if logo_path and os.path.exists(logo_path):
        lp = _tight_para(doc, space_after=2)
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        lp.add_run().add_picture(logo_path, height=Cm(2.5))
    tp = doc.add_heading(f"정답 및 해설 ({title})", 1)
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_after = Pt(4)

    new_sec = doc.add_section(WD_SECTION.CONTINUOUS)
    _set_two_columns(new_sec)

    ans_matches = re.findall(r"\b(\d+)(?:번|[\)\.])?(?:\s*정답\s*:?)?\s*[\(\[]?([①②③④⑤])[\)\]]?", raw_ans)
    if ans_matches:
        p_title = _tight_para(doc, space_before=2, space_after=3)
        r = p_title.add_run("[ OMR 정답 카드 ]")
        r.bold = True
        r.font.size = Pt(12)
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        tbl = doc.add_table(rows=1, cols=10)
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        for i in range(5):
            hdr[i * 2].text = "번호"
            hdr[i * 2 + 1].text = "정답"
            for c in (hdr[i * 2], hdr[i * 2 + 1]):
                _set_cell_bg(c, "2563EB")
                _tight_table_cell(c)
                for p in c.paragraphs:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if p.runs:
                        p.runs[0].bold = True
                        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)

        row_cells = None
        for i, (qn_str, ans) in enumerate(ans_matches):
            if i % 5 == 0:
                row_cells = tbl.add_row().cells
            row_cells[(i % 5) * 2].text = qn_str
            row_cells[(i % 5) * 2 + 1].text = ans
            _set_cell_bg(row_cells[(i % 5) * 2], "EFF6FF")
            for p in (row_cells[(i % 5) * 2].paragraphs[0], row_cells[(i % 5) * 2 + 1].paragraphs[0]):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _tight_para(doc, space_before=4, space_after=4)

    for line in raw_ans.split("\n"):
        ls = line.strip()
        if ls:
            p = _tight_para(doc, space_before=1, space_after=1)
            p.add_run(ls)

    dx, px = build_paths(target_dir, filename, prefix="답안지_")
    doc.save(dx)
    return dx, px

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 5. 문제 은행 (JSON) — 정답/해설 포함 저장
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def save_to_qbank(questions, tag="", answers_map=None):
    """파싱된 문제들을 JSON 문제 은행에 추가.
    answers_map: {원본번호: 해설텍스트} 딕셔너리 (optional)
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
        # 해당 문항의 정답/해설 연결
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


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 6. 생성 히스토리
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _add_history(entry):
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


def _load_history():
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 7. 자동 임시저장
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _save_draft(exam_text, ans_text, title):
    try:
        with open(DRAFT_FILE, "w", encoding="utf-8") as f:
            json.dump({"exam": exam_text, "answer": ans_text, "title": title,
                        "saved_at": datetime.datetime.now().isoformat()}, f, ensure_ascii=False)
    except OSError:
        pass


def _load_draft():
    if not os.path.exists(DRAFT_FILE):
        return None
    try:
        with open(DRAFT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


def _clear_draft():
    if os.path.exists(DRAFT_FILE):
        try:
            os.remove(DRAFT_FILE)
        except OSError:
            pass

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 8. 메인 앱
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
class QuestionBankWindow(tk.Toplevel):
    """문제은행 통합 관리 브라우저 — 다중 필터 검색, 정답/해설 편집, 메타데이터(난이도/오답률) 수정, 시험지 생성."""

    def __init__(self, master, db_root, on_generate_callback=None):
        super().__init__(master)
        self.title("문제은행 데이터베이스 (다중 필터 + 메타데이터)")
        self.geometry("1300x850")
        self.db_root = db_root
        self.on_generate = on_generate_callback
        self.current_item = None
        self.answer_vars = []
        self.explanation_widgets = []
        self._editor_widgets = []
        self._meta_vars = []
        self._build_ui()
        self._refresh_list()

    # ── UI 구축 ──
    def _build_ui(self):
        # 상단 검색바 (다중 필터)
        top = tk.Frame(self)
        top.pack(fill="x", padx=10, pady=6)

        tk.Label(top, text="검색 대상:", font=("맑은 고딕", 9)).pack(side="left")
        self.cb_field = ttk.Combobox(top, values=["전체", "태그", "본문"], width=8, state="readonly")
        self.cb_field.set("전체")
        self.cb_field.pack(side="left", padx=2)

        tk.Label(top, text="정답 유무:", font=("맑은 고딕", 9)).pack(side="left", padx=(10, 0))
        self.cb_ans_filter = ttk.Combobox(top, values=["전체", "정답 존재(O)", "정답 없음(X)"],
                                          width=12, state="readonly")
        self.cb_ans_filter.set("전체")
        self.cb_ans_filter.pack(side="left", padx=2)

        tk.Label(top, text="검색어:", font=("맑은 고딕", 9)).pack(side="left", padx=(10, 0))
        self.ent_search = tk.Entry(top, width=25, font=("맑은 고딕", 10))
        self.ent_search.pack(side="left", padx=4)
        self.ent_search.bind("<Return>", lambda e: self._do_search())
        tk.Button(top, text="적용", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._do_search).pack(side="left", padx=2)
        tk.Button(top, text="전체 보기", command=lambda: self._refresh_list()).pack(side="left", padx=2)
        tk.Button(top, text="인덱스 재생성", bg="#6366f1", fg="white",
                  font=("맑은 고딕", 9), command=self._rebuild_index).pack(side="left", padx=6)
        self.lbl_count = tk.Label(top, text="", font=("맑은 고딕", 9), fg="#64748b")
        self.lbl_count.pack(side="right", padx=8)

        # PanedWindow — 좌(리스트) / 우(편집기)
        pw = tk.PanedWindow(self, orient=tk.HORIZONTAL, sashwidth=6)
        pw.pack(fill="both", expand=True, padx=10, pady=(0, 6))

        # 좌: Treeview
        left = tk.Frame(pw)
        pw.add(left, width=560, minsize=300)
        cols = ("ID", "제목", "태그", "정답", "난이도", "수정일", "문항수")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", height=28, selectmode="extended")
        for c, w in zip(cols, [90, 120, 90, 45, 45, 85, 45]):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, minwidth=35)
        vsb = tk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        # 우: 편집 영역
        right = tk.Frame(pw)
        pw.add(right, minsize=400)

        # 제시문 (passage)
        lf_pass = tk.LabelFrame(right, text=" 제시문 (Passage) ", font=("맑은 고딕", 9, "bold"))
        lf_pass.pack(fill="x", padx=6, pady=(4, 2))
        self.txt_passage = tk.Text(lf_pass, height=4, font=("Consolas", 9), bg="#f8fafc", state="disabled", wrap=tk.WORD)
        self.txt_passage.pack(fill="x", padx=4, pady=4)

        # 문항 편집 (스크롤 가능)
        lf_q = tk.LabelFrame(right, text=" 문항 편집 ", font=("맑은 고딕", 9, "bold"))
        lf_q.pack(fill="both", expand=True, padx=6, pady=2)

        canvas = tk.Canvas(lf_q, bg="#ffffff", highlightthickness=0)
        vsb2 = tk.Scrollbar(lf_q, orient="vertical", command=canvas.yview)
        self.q_frame = tk.Frame(canvas, bg="#ffffff")
        self.q_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.q_frame, anchor="nw")
        canvas.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side="right", fill="y")
        canvas.pack(fill="both", expand=True)
        # 마우스 휠 스크롤
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        # 태그 표시
        tag_row = tk.Frame(right)
        tag_row.pack(fill="x", padx=6, pady=2)
        tk.Label(tag_row, text="태그:", font=("맑은 고딕", 9, "bold")).pack(side="left")
        self.lbl_tags = tk.Label(tag_row, text="", font=("맑은 고딕", 9), fg="#059669")
        self.lbl_tags.pack(side="left", padx=4)

        # 하단 버튼
        bottom = tk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=8)
        tk.Button(bottom, text="저장", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=10, command=self._save_current).pack(side="left", padx=4)
        tk.Button(bottom, text="태그 수정", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=10, command=self._edit_tags).pack(side="left", padx=4)
        tk.Button(bottom, text="삭제", bg="#dc2626", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=10, command=self._delete_selected).pack(side="left", padx=4)
        tk.Button(bottom, text="시험지 생성", bg="#ea580c", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=14, command=self._generate_from_selected).pack(side="right", padx=4)
        tk.Button(bottom, text="JSON Export", bg="#0891b2", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=14, command=self._export_selected).pack(side="right", padx=4)

    # ── 리스트 (다중 필터 지원) ──
    def _refresh_list(self, query=""):
        for c in self.tree.get_children():
            self.tree.delete(c)
        index = exam_db.load_index(self.db_root)
        field = self.cb_field.get() if hasattr(self, "cb_field") else "전체"
        ans_filter = self.cb_ans_filter.get() if hasattr(self, "cb_ans_filter") else "전체"
        query_lower = query.lower() if query else ""

        count = 0
        for iid, entry in index.get("items", {}).items():
            # 정답 유무 필터
            has_answer = entry.get("has_answer", False)
            if ans_filter == "정답 존재(O)" and not has_answer:
                continue
            if ans_filter == "정답 없음(X)" and has_answer:
                continue

            # 검색어 필터 (대상별)
            if query_lower:
                if field == "태그":
                    search_str = " ".join(entry.get("tags", [])).lower()
                elif field == "본문":
                    search_str = entry.get("passage", "").lower()
                else:  # 전체
                    search_str = entry.get("search_text", "")
                if query_lower not in search_str:
                    continue

            self.tree.insert("", "end", iid=iid, values=(
                iid,
                entry.get("source_title", "")[:25],
                ", ".join(entry.get("tags", []))[:20],
                "O" if has_answer else "X",
                entry.get("difficulty", "3"),
                entry.get("updated_at", "")[:10],
                entry.get("question_count", 0),
            ))
            count += 1
        self.lbl_count.config(text=f"결과: {count}건")

    def _do_search(self):
        q = self.ent_search.get().strip()
        self._refresh_list(q)

    # ── 선택 → 편집기 로드 ──
    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        item = exam_db.load_item(self.db_root, item_id)
        if not item:
            return
        self.current_item = item
        self._load_editor(item)

    def _load_editor(self, item):
        # 제시문
        self.txt_passage.config(state="normal")
        self.txt_passage.delete("1.0", tk.END)
        self.txt_passage.insert("1.0", item.get("passage", "") or "(없음)")
        self.txt_passage.config(state="disabled")
        # 태그
        self.lbl_tags.config(text=", ".join(item.get("tags", [])) or "(없음)")

        # 기존 위젯 정리
        for w in self._editor_widgets:
            w.destroy()
        self._editor_widgets.clear()
        self.answer_vars.clear()
        self.explanation_widgets.clear()
        self._meta_vars.clear()

        questions = item.get("questions", [])
        for idx, q in enumerate(questions):
            c_type = q.get("c_type", "num")
            frame = tk.Frame(self.q_frame, bg="#ffffff", bd=1, relief="groove")
            frame.pack(fill="x", padx=4, pady=4)
            self._editor_widgets.append(frame)

            # 문항 헤더
            hdr = tk.Label(frame, text=f"문항 {idx+1}  (원본#{q.get('num', '?')})",
                           font=("맑은 고딕", 9, "bold"), bg="#eff6ff", anchor="w")
            hdr.pack(fill="x", padx=4, pady=(4, 2))

            # 본문
            txt_lbl = tk.Label(frame, text=q.get("text", "")[:200],
                               font=("맑은 고딕", 9), wraplength=550, justify="left", bg="#ffffff", anchor="w")
            txt_lbl.pack(fill="x", padx=8, pady=2)

            # 선지
            choices = q.get("choices", [])
            for ci, ch in enumerate(choices):
                if ch:
                    mark = CIRCLE_NUMS[ci] if c_type == "num" else ALPHA_CHOICES[ci] if ci < 5 else ""
                    clbl = tk.Label(frame, text=f"  {mark} {ch}",
                                    font=("맑은 고딕", 8), bg="#ffffff", anchor="w", fg="#475569")
                    clbl.pack(fill="x", padx=12)

            # 정답 + 메타데이터 행
            ans_frame = tk.Frame(frame, bg="#ffffff")
            ans_frame.pack(fill="x", padx=8, pady=2)
            tk.Label(ans_frame, text="정답:", font=("맑은 고딕", 9, "bold"), bg="#ffffff").pack(side="left")
            if c_type == "alpha":
                ans_values = ["", "A", "B", "C", "D", "E"]
            else:
                ans_values = ["", "①", "②", "③", "④", "⑤"]
            ans_var = tk.StringVar(value=q.get("answer") or "")
            ans_cb = ttk.Combobox(ans_frame, textvariable=ans_var, values=ans_values,
                                  width=6, state="readonly")
            ans_cb.pack(side="left", padx=4)
            self.answer_vars.append(ans_var)

            # 메타데이터: 난이도 + 오답률
            tk.Label(ans_frame, text="난이도(1~5):", font=("맑은 고딕", 8, "bold"),
                     bg="#ffffff").pack(side="left", padx=(12, 0))
            diff_var = tk.StringVar(value=str(q.get("difficulty", "3")))
            tk.Spinbox(ans_frame, from_=1, to=5, textvariable=diff_var, width=3,
                       font=("맑은 고딕", 9)).pack(side="left", padx=2)

            tk.Label(ans_frame, text="오답률(%):", font=("맑은 고딕", 8, "bold"),
                     bg="#ffffff").pack(side="left", padx=(8, 0))
            err_var = tk.StringVar(value=str(q.get("error_rate", "0")))
            tk.Entry(ans_frame, textvariable=err_var, width=5,
                     font=("맑은 고딕", 9)).pack(side="left", padx=2)

            self._meta_vars.append({"diff": diff_var, "err": err_var})

            # 해설 Text
            tk.Label(ans_frame, text="해설:", font=("맑은 고딕", 9, "bold"), bg="#ffffff").pack(side="left", padx=(12, 0))
            expl_txt = tk.Text(frame, height=2, font=("맑은 고딕", 9), wrap=tk.WORD)
            expl_txt.pack(fill="x", padx=8, pady=(0, 4))
            expl_txt.insert("1.0", q.get("explanation") or "")
            self.explanation_widgets.append(expl_txt)

    # ── 저장 (정답 + 해설 + 메타데이터) ──
    def _save_current(self):
        if not self.current_item:
            return messagebox.showwarning("알림", "편집할 아이템을 선택하세요.", parent=self)
        item = self.current_item
        questions = item.get("questions", [])
        for idx, q in enumerate(questions):
            if idx < len(self.answer_vars):
                raw_ans = self.answer_vars[idx].get()
                q["answer"] = exam_db.validate_answer(raw_ans, q.get("c_type", "num"))
                if raw_ans and q["answer"] is None and raw_ans.strip():
                    messagebox.showwarning("정답 오류",
                                           f"문항 {idx+1}: '{raw_ans}'은(는) 유효하지 않은 정답입니다.",
                                           parent=self)
            if idx < len(self.explanation_widgets):
                q["explanation"] = self.explanation_widgets[idx].get("1.0", tk.END).strip() or None
            # 메타데이터 저장
            if idx < len(self._meta_vars):
                q["difficulty"] = self._meta_vars[idx]["diff"].get()
                q["error_rate"] = self._meta_vars[idx]["err"].get()
        exam_db.save_item(self.db_root, item)
        exam_db.update_index_entry(self.db_root, item)
        self._refresh_list(self.ent_search.get().strip())
        messagebox.showinfo("저장", "정답/해설 및 메타데이터가 저장되었습니다.", parent=self)

    # ── 태그 수정 ──
    def _edit_tags(self):
        if not self.current_item:
            return messagebox.showwarning("알림", "아이템을 선택하세요.", parent=self)
        current = ", ".join(self.current_item.get("tags", []))
        new_tags = simpledialog.askstring("태그 수정", f"태그 (쉼표 구분):\n현재: {current}",
                                          initialvalue=current, parent=self)
        if new_tags is None:
            return
        self.current_item["tags"] = [t.strip() for t in new_tags.split(",") if t.strip()]
        exam_db.save_item(self.db_root, self.current_item)
        exam_db.update_index_entry(self.db_root, self.current_item)
        self.lbl_tags.config(text=", ".join(self.current_item["tags"]) or "(없음)")
        self._refresh_list(self.ent_search.get().strip())

    # ── 삭제 ──
    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("알림", "삭제할 아이템을 선택하세요.", parent=self)
        if not messagebox.askyesno("확인", f"{len(sel)}개 아이템을 삭제하시겠습니까?", parent=self):
            return
        for iid in sel:
            exam_db.delete_item(self.db_root, iid)
            exam_db.remove_index_entry(self.db_root, iid)
        self.current_item = None
        self._refresh_list(self.ent_search.get().strip())

    # ── 시험지 생성 (선택한 아이템으로) ──
    def _generate_from_selected(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("알림", "시험지를 생성할 아이템을 선택하세요.", parent=self)
        items = [exam_db.load_item(self.db_root, iid) for iid in sel]
        items = [it for it in items if it]
        if not items:
            return
        if self.on_generate:
            self.on_generate(items)
            self.destroy()

    # ── JSON Export ──
    def _export_selected(self):
        sel = self.tree.selection()
        if not sel:
            return messagebox.showwarning("알림", "Export할 아이템을 선택하세요.", parent=self)
        items = [exam_db.load_item(self.db_root, iid) for iid in sel]
        items = [it for it in items if it]
        if not items:
            return
        target_dir = filedialog.askdirectory(title="Export 저장 폴더", parent=self)
        if not target_dir:
            return
        title = simpledialog.askstring("제목", "Export 파일명:", initialvalue="exam_export", parent=self)
        if not title:
            return
        path = exam_db.export_from_db_items(items, target_dir, title,
                                             display_title=title)
        messagebox.showinfo("Export 완료", f"저장: {path}", parent=self)

    # ── 인덱스 재생성 ──
    def _rebuild_index(self):
        idx = exam_db.build_index(self.db_root)
        cnt = len(idx.get("items", {}))
        self._refresh_list(self.ent_search.get().strip())
        messagebox.showinfo("완료", f"인덱스 재생성: {cnt}개 아이템", parent=self)


class SutamMakerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("사탐/경제 모의고사 자동 생성기 PRO v4.0")
        self.root.geometry("1100x980")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.cfg = load_config()
        self._build_ui()
        self._bind_shortcuts()
        self._restore_draft()
        self._start_autosave()
        # Drag & Drop 바인딩 (tkinterdnd2 설치 시)
        if TkinterDnD and hasattr(self.root, "drop_target_register"):
            try:
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind("<<Drop>>", self._on_file_drop)
            except Exception:
                pass

    # --- UI 구축 ---
    def _build_ui(self):
        # 상태바
        self.status_bar = tk.Frame(self.root, bg="#e2e8f0", height=26)
        self.status_bar.pack(fill="x", side="bottom")
        self.status_bar.pack_propagate(False)
        self.lbl_status = tk.Label(self.status_bar, text="", font=("맑은 고딕", 8), bg="#e2e8f0", fg="#475569")
        self.lbl_status.pack(side="left", padx=10)

        # 메인 탭
        self.tabs = ttk.Notebook(self.root)
        self.tabs.pack(fill="both", expand=True, padx=8, pady=(8, 4))

        self.tab_main = ttk.Frame(self.tabs)
        self.tab_preview = ttk.Frame(self.tabs)
        self.tab_qbank = ttk.Frame(self.tabs)
        self.tab_history = ttk.Frame(self.tabs)
        self.tab_ai = ttk.Frame(self.tabs)
        self.tab_db = ttk.Frame(self.tabs)

        self.tabs.add(self.tab_main, text="  시험지 생성  ")
        self.tabs.add(self.tab_preview, text="  미리보기  ")
        self.tabs.add(self.tab_qbank, text="  문제 은행  ")
        self.tabs.add(self.tab_history, text="  생성 히스토리  ")
        self.tabs.add(self.tab_ai, text="  AI 명령어  ")
        self.tabs.add(self.tab_db, text="  DB 도구  ")

        self._build_main_tab()
        self._build_preview_tab()
        self._build_qbank_tab()
        self._build_history_tab()
        self._build_ai_tab()
        self._build_db_tab()
        self._update_status()

    def _bind_shortcuts(self):
        self.root.bind("<Control-g>", lambda e: self._generate("both"))
        self.root.bind("<Control-p>", lambda e: self._show_preview())
        self.root.bind("<Control-s>", lambda e: self._save_draft_now())

    # --- TAB 1: 시험지 생성 ---
    def _build_main_tab(self):
        f_cfg = tk.LabelFrame(self.tab_main, text=" 설정 ", font=("맑은 고딕", 10, "bold"))
        f_cfg.pack(fill="x", padx=12, pady=6)

        r0 = tk.Frame(f_cfg)
        r0.pack(fill="x", padx=8, pady=3)
        tk.Label(r0, text="저장 폴더:", width=12, anchor="w").pack(side="left")
        self.dir_var = tk.StringVar(value=self.cfg.get("last_dir", SCRIPT_DIR))
        tk.Entry(r0, textvariable=self.dir_var, width=42).pack(side="left", padx=4)
        tk.Button(r0, text="찾아보기", command=self._browse_dir).pack(side="left", padx=2)
        tk.Button(r0, text="폴더 열기", command=self._open_target_dir).pack(side="left", padx=2)

        r1 = tk.Frame(f_cfg)
        r1.pack(fill="x", padx=8, pady=3)
        tk.Label(r1, text="배너 로고:", width=12, anchor="w").pack(side="left")
        self.logo_var = tk.StringVar(value=self.cfg.get("last_logo", ""))
        tk.Entry(r1, textvariable=self.logo_var, width=42).pack(side="left", padx=4)
        tk.Button(r1, text="이미지 선택", command=self._browse_logo).pack(side="left", padx=2)
        tk.Button(r1, text="초기화", command=lambda: self.logo_var.set("")).pack(side="left", padx=2)

        r2 = tk.Frame(f_cfg)
        r2.pack(fill="x", padx=8, pady=3)
        tk.Label(r2, text="파일 제목:", width=12, anchor="w").pack(side="left")
        self.ent_title = tk.Entry(r2, width=30)
        self.ent_title.pack(side="left", padx=4)
        tk.Label(r2, text="(예: 이준호_경제)", fg="gray", font=("맑은 고딕", 8)).pack(side="left")

        r3 = tk.Frame(f_cfg)
        r3.pack(fill="x", padx=8, pady=3)
        tk.Label(r3, text="폰트:", width=12, anchor="w").pack(side="left")
        self.font_combo = ttk.Combobox(r3, values=["맑은 고딕", "바탕", "굴림", "나눔고딕"], width=14, state="readonly")
        self.font_combo.set(self.cfg.get("font_name", "맑은 고딕"))
        self.font_combo.pack(side="left", padx=4)
        tk.Label(r3, text="크기:").pack(side="left")
        self.size_combo = ttk.Combobox(r3, values=["9", "10", "11", "12"], width=4, state="readonly")
        self.size_combo.set(self.cfg.get("font_size", "10"))
        self.size_combo.pack(side="left", padx=4)

        r4 = tk.Frame(f_cfg)
        r4.pack(fill="x", padx=8, pady=3)
        tk.Label(r4, text="배치 생성:", width=12, anchor="w").pack(side="left")
        self.ent_batch = tk.Entry(r4, width=30)
        self.ent_batch.pack(side="left", padx=4)
        tk.Label(r4, text="(쉼표로 구분: 이준호,김철수,박영희)", fg="gray", font=("맑은 고딕", 8)).pack(side="left")

        r5 = tk.Frame(f_cfg)
        r5.pack(fill="x", padx=8, pady=3)
        self.var_shuffle = tk.BooleanVar(value=False)
        tk.Checkbutton(r5, text="문제 순서 섞기 (학생마다 랜덤)", variable=self.var_shuffle).pack(side="left", padx=(88, 10))
        tk.Label(r5, text="워터마크:").pack(side="left")
        self.ent_watermark = tk.Entry(r5, width=15)
        self.ent_watermark.pack(side="left", padx=4)

        # 문제 입력
        tk.Label(self.tab_main, text="[1] 문제 입력 (AI 출력물 붙여넣기 가능)",
                 font=("맑은 고딕", 10, "bold"), fg="#2563eb").pack(anchor="w", padx=12, pady=(8, 0))
        self.txt_exam = tk.Text(self.tab_main, wrap=tk.WORD, height=14, font=("Consolas", 10))
        self.txt_exam.pack(fill="both", expand=True, padx=12, pady=4)
        self.txt_exam.bind("<KeyRelease>", self._on_text_change)

        # 정답 입력
        tk.Label(self.tab_main, text="[2] 정답 및 해설 (선택 - AI가 생성한 정답/해설 붙여넣기)",
                 font=("맑은 고딕", 10, "bold"), fg="#16a34a").pack(anchor="w", padx=12, pady=(4, 0))
        self.txt_ans = tk.Text(self.tab_main, wrap=tk.WORD, height=7, font=("Consolas", 10))
        self.txt_ans.pack(fill="x", padx=12, pady=4)

        # 생성 버튼
        bf = tk.Frame(self.tab_main)
        bf.pack(pady=8)
        tk.Button(bf, text="Word만", bg="#0891b2", fg="white", font=("맑은 고딕", 10, "bold"),
                  width=14, height=2, command=lambda: self._generate("word")).pack(side="left", padx=6)
        tk.Button(bf, text="PDF만", bg="#dc2626", fg="white", font=("맑은 고딕", 10, "bold"),
                  width=14, height=2, command=lambda: self._generate("pdf")).pack(side="left", padx=6)
        tk.Button(bf, text="Word + PDF", bg="#ea580c", fg="white", font=("맑은 고딕", 10, "bold"),
                  width=18, height=2, command=lambda: self._generate("both")).pack(side="left", padx=6)
        tk.Button(bf, text="미리보기 (Ctrl+P)", bg="#7c3aed", fg="white", font=("맑은 고딕", 10, "bold"),
                  width=18, height=2, command=self._show_preview).pack(side="left", padx=6)

    # --- TAB 2: 미리보기 ---
    def _build_preview_tab(self):
        ctrl = tk.Frame(self.tab_preview)
        ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="파싱 새로고침", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._show_preview).pack(side="left", padx=4)
        tk.Button(ctrl, text="문제 은행에 저장 (정답 포함)", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._save_to_bank_from_preview).pack(side="left", padx=4)
        self.lbl_preview_count = tk.Label(ctrl, text="", font=("맑은 고딕", 9), fg="#64748b")
        self.lbl_preview_count.pack(side="right", padx=8)

        self.txt_preview = tk.Text(self.tab_preview, font=("Consolas", 10), bg="#f8fafc", state="disabled")
        scroll = tk.Scrollbar(self.tab_preview, command=self.txt_preview.yview)
        self.txt_preview.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.txt_preview.pack(fill="both", expand=True, padx=12, pady=6)

    def _show_preview(self, event=None):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showwarning("경고", "문제 텍스트를 먼저 입력하세요.")
            return
        parsed = parse_exam_text(raw)
        self.txt_preview.config(state="normal")
        self.txt_preview.delete("1.0", tk.END)
        if not parsed:
            first_lines = raw[:300].replace('\r', '\\r')
            debug_msg = (
                "파싱된 문제가 없습니다.\n\n"
                "── 지원 형식 ──\n"
                "  01 문제 본문...\n"
                "  01. 문제 본문...\n"
                "  1) 문제 본문...\n\n"
                "── 입력된 텍스트 (처음 300자) ──\n"
                f"{first_lines}\n"
            )
            self.txt_preview.insert(tk.END, debug_msg)
            self.txt_preview.config(state="disabled")
            return
        for i, q in enumerate(parsed, 1):
            self.txt_preview.insert(tk.END, f"{'='*50}\n")
            self.txt_preview.insert(tk.END, f"  문항 {i} (원본 번호: {q['num']})\n")
            self.txt_preview.insert(tk.END, f"{'='*50}\n\n")
            self.txt_preview.insert(tk.END, f"[본문]\n{q['text']}\n\n")
            if q["jesi"]:
                self.txt_preview.insert(tk.END, f"[제시문]\n{q['jesi']}\n\n")
            if q["bogi"]:
                self.txt_preview.insert(tk.END, f"[보기]\n{q['bogi']}\n\n")
            if q["tables"]:
                self.txt_preview.insert(tk.END, f"[표] {len(q['tables'])}개 감지됨\n\n")
            if q.get("image_path"):
                self.txt_preview.insert(tk.END, f"[이미지] {q['image_path']}\n\n")
            if q.get("graph_tag"):
                self.txt_preview.insert(tk.END, f"[그래프 태그]\n{q['graph_tag']}\n\n")
            if q["choices"]:
                self.txt_preview.insert(tk.END, "[선택지]\n")
                for j, ch in enumerate(q["choices"]):
                    if ch:
                        self.txt_preview.insert(tk.END, f"  {CIRCLE_NUMS[j]} {ch}\n")
                self.txt_preview.insert(tk.END, "\n")
        self.txt_preview.config(state="disabled")
        self.lbl_preview_count.config(text=f"총 {len(parsed)}문항 파싱됨")
        self.tabs.select(self.tab_preview)
        self._set_status(f"미리보기: {len(parsed)}문항 파싱 완료")

    def _save_to_bank_from_preview(self):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        if not raw:
            return messagebox.showwarning("경고", "문제 텍스트가 없습니다.")
        parsed = parse_exam_text(raw)
        if not parsed:
            return messagebox.showwarning("오류", "파싱된 문제가 없습니다.")
        # 정답/해설도 함께 저장
        raw_ans = self.txt_ans.get("1.0", tk.END).strip()
        answers_map = parse_answer_by_question(raw_ans)
        tag = simpledialog.askstring("태그", "저장 태그 (예: 경제_중간고사):", parent=self.root)
        if tag is None:
            return
        cnt = save_to_qbank(parsed, tag=tag or "", answers_map=answers_map)
        messagebox.showinfo("완료", f"{cnt}문항을 문제 은행에 저장했습니다.\n(정답/해설 {'포함' if answers_map else '없음'})")
        self._refresh_qbank()

    # --- TAB 3: 문제 은행 (쇼핑 모드 + 정답 매칭) ---
    def _build_qbank_tab(self):
        ctrl = tk.Frame(self.tab_qbank)
        ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="새로고침", command=self._refresh_qbank).pack(side="left", padx=4)
        tk.Button(ctrl, text="선택 문항 -> 시험지+답안지", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._load_from_bank).pack(side="left", padx=4)
        tk.Button(ctrl, text="전체 선택", bg="#6366f1", fg="white",
                  font=("맑은 고딕", 9), command=self._select_all_bank).pack(side="left", padx=4)
        tk.Button(ctrl, text="선택 해제", bg="#94a3b8", fg="white",
                  font=("맑은 고딕", 9), command=self._deselect_all_bank).pack(side="left", padx=4)
        tk.Button(ctrl, text="전체 삭제", bg="#dc2626", fg="white",
                  command=self._clear_bank).pack(side="right", padx=4)
        tk.Button(ctrl, text="선택 항목 삭제", bg="#f97316", fg="white",
                  command=self._delete_selected_bank).pack(side="right", padx=4)
        self.lbl_bank_count = tk.Label(ctrl, text="", font=("맑은 고딕", 9), fg="#64748b")
        self.lbl_bank_count.pack(side="right", padx=8)

        # 안내 라벨
        info = tk.Label(self.tab_qbank,
                        text="서로 다른 세트에서 자유롭게 골라 조합하세요. 번호는 자동으로 1번부터 재매김됩니다. 정답/해설도 자동 매칭!",
                        font=("맑은 고딕", 9), fg="#059669", bg="#ecfdf5", relief="groove", padx=8, pady=4)
        info.pack(fill="x", padx=12, pady=(0, 4))

        cols = ("번호", "태그", "정답유무", "저장일", "내용 미리보기")
        self.tree_bank = ttk.Treeview(self.tab_qbank, columns=cols, show="headings", height=18, selectmode="extended")
        for c, w in zip(cols, [50, 100, 70, 110, 450]):
            self.tree_bank.heading(c, text=c)
            self.tree_bank.column(c, width=w)
        self.tree_bank.pack(fill="both", expand=True, padx=12, pady=6)

        # 하단: 선택한 문항 미리보기
        self.lbl_selected = tk.Label(self.tab_qbank, text="선택된 문항: 0개",
                                     font=("맑은 고딕", 10, "bold"), fg="#2563eb")
        self.lbl_selected.pack(anchor="w", padx=12, pady=(4, 2))
        self.tree_bank.bind("<<TreeviewSelect>>", self._on_bank_select)

        self._refresh_qbank()

    def _on_bank_select(self, event=None):
        sel = self.tree_bank.selection()
        self.lbl_selected.config(text=f"선택된 문항: {len(sel)}개")

    def _select_all_bank(self):
        all_items = self.tree_bank.get_children()
        self.tree_bank.selection_set(all_items)
        self._on_bank_select()

    def _deselect_all_bank(self):
        self.tree_bank.selection_remove(self.tree_bank.selection())
        self._on_bank_select()

    def _refresh_qbank(self):
        for i in self.tree_bank.get_children():
            self.tree_bank.delete(i)
        bank = load_qbank()
        for i, q in enumerate(bank, 1):
            preview = q.get("text", "")[:60].replace("\n", " ")
            saved = q.get("saved_at", "")[:10]
            tag = q.get("tag", "")
            has_ans = "O" if q.get("answer", "").strip() else "X"
            self.tree_bank.insert("", "end", iid=str(i - 1), values=(i, tag, has_ans, saved, preview))
        self.lbl_bank_count.config(text=f"총 {len(bank)}문항")

    def _load_from_bank(self):
        """선택한 문항들을 시험지+답안지에 동시 로드. 번호를 1부터 재매김."""
        sel = self.tree_bank.selection()
        if not sel:
            return messagebox.showwarning("알림", "문항을 선택하세요. (Ctrl+클릭으로 여러 개 선택 가능)")
        bank = load_qbank()
        exam_lines = []
        ans_lines = []
        new_num = 1

        for iid in sel:
            idx = int(iid)
            if idx >= len(bank):
                continue
            q = bank[idx]

            # --- 시험지 텍스트 재구성 ---
            body = f"{new_num}. {q['text']}"
            if q.get("jesi"):
                body += f"\n<제시문>{q['jesi']}</제시문>"
            if q.get("graph_tag"):
                body += f"\n{q['graph_tag']}"
            if q.get("bogi"):
                body += f"\n<보기>\n{q['bogi']}"
            if q.get("choices"):
                for j, ch in enumerate(q["choices"]):
                    if ch:
                        body += f"\n{CIRCLE_NUMS[j]} {ch}"
            exam_lines.append(body)

            # --- 정답/해설 텍스트 재구성 (번호 재매김) ---
            ans_raw = q.get("answer", "").strip()
            if ans_raw:
                # 원본 번호를 새 번호로 교체
                old_num = q.get("num", "")
                renamed = re.sub(
                    rf"^\s*{re.escape(old_num)}\s*(?:번|[\.\)])",
                    f"{new_num}번",
                    ans_raw,
                    count=1,
                    flags=re.MULTILINE
                )
                # 정답 라인의 번호도 교체
                renamed = re.sub(
                    rf"\b{re.escape(old_num)}\s*(?:번\s*)?정답",
                    f"{new_num}번 정답",
                    renamed,
                    flags=re.IGNORECASE
                )
                ans_lines.append(renamed)
            else:
                # 정답 정보가 없으면 빈 라인 표시
                ans_lines.append(f"{new_num}번 정답: (정보 없음)")

            new_num += 1

        exam_text = "\n\n".join(exam_lines)
        ans_text = "\n\n".join(ans_lines)

        # 기존 내용 처리
        current_exam = self.txt_exam.get("1.0", tk.END).strip()
        current_ans = self.txt_ans.get("1.0", tk.END).strip()

        if current_exam:
            choice = messagebox.askyesnocancel(
                "불러오기 방법",
                f"선택한 {len(sel)}문항을 불러옵니다.\n\n"
                "[예] 기존 내용 뒤에 추가\n"
                "[아니오] 기존 내용 지우고 새로 입력\n"
                "[취소] 취소"
            )
            if choice is None:
                return
            elif choice:
                self.txt_exam.insert(tk.END, "\n\n" + exam_text)
                self.txt_ans.insert(tk.END, "\n\n" + ans_text)
            else:
                self.txt_exam.delete("1.0", tk.END)
                self.txt_exam.insert("1.0", exam_text)
                self.txt_ans.delete("1.0", tk.END)
                self.txt_ans.insert("1.0", ans_text)
        else:
            self.txt_exam.insert("1.0", exam_text)
            self.txt_ans.insert("1.0", ans_text)

        self.tabs.select(self.tab_main)
        self._set_status(f"문제 은행에서 {len(sel)}문항 + 정답/해설 불러옴")

    def _delete_selected_bank(self):
        sel = self.tree_bank.selection()
        if not sel:
            return messagebox.showwarning("알림", "삭제할 문항을 선택하세요.")
        if not messagebox.askyesno("확인", f"선택한 {len(sel)}개 문항을 삭제하시겠습니까?"):
            return
        bank = load_qbank()
        indices = sorted([int(iid) for iid in sel], reverse=True)
        for idx in indices:
            if idx < len(bank):
                bank.pop(idx)
        with open(QBANK_FILE, "w", encoding="utf-8") as f:
            json.dump(bank, f, ensure_ascii=False, indent=2)
        self._refresh_qbank()

    def _clear_bank(self):
        if not messagebox.askyesno("확인", "문제 은행을 전체 삭제하시겠습니까?"):
            return
        clear_qbank()
        self._refresh_qbank()

    # --- TAB 4: 생성 히스토리 ---
    def _build_history_tab(self):
        ctrl = tk.Frame(self.tab_history)
        ctrl.pack(fill="x", padx=12, pady=6)
        tk.Button(ctrl, text="새로고침", command=self._refresh_history).pack(side="left", padx=4)
        tk.Button(ctrl, text="폴더 열기", command=self._open_history_dir).pack(side="left", padx=4)

        cols = ("시각", "제목", "학생", "문항수", "형식", "경로")
        self.tree_hist = ttk.Treeview(self.tab_history, columns=cols, show="headings", height=18)
        for c, w in zip(cols, [130, 140, 100, 60, 70, 350]):
            self.tree_hist.heading(c, text=c)
            self.tree_hist.column(c, width=w)
        self.tree_hist.pack(fill="both", expand=True, padx=12, pady=6)
        self._refresh_history()

    def _refresh_history(self):
        for i in self.tree_hist.get_children():
            self.tree_hist.delete(i)
        for h in _load_history():
            self.tree_hist.insert("", "end", values=(
                h.get("timestamp", "")[:19].replace("T", " "),
                h.get("title", ""),
                h.get("student", ""),
                h.get("q_count", ""),
                h.get("format", ""),
                h.get("path", ""),
            ))

    def _open_history_dir(self):
        sel = self.tree_hist.selection()
        if sel:
            vals = self.tree_hist.item(sel[0])["values"]
            path = vals[-1] if vals else ""
            if path and os.path.isdir(str(path)):
                self._open_dir(str(path))
                return
        self._open_dir(self.dir_var.get())

    # --- TAB 5: AI 명령어 ---
    def _build_ai_tab(self):
        # 상단 설정
        f_settings = tk.LabelFrame(self.tab_ai, text=" AI 프롬프트 설정 ", font=("맑은 고딕", 10, "bold"))
        f_settings.pack(fill="x", padx=12, pady=6)

        r0 = tk.Frame(f_settings)
        r0.pack(fill="x", padx=8, pady=3)
        tk.Label(r0, text="과목:", width=10, anchor="w").pack(side="left")
        self.ai_subject = ttk.Combobox(r0, values=SUBJECT_LIST, width=20, state="readonly")
        self.ai_subject.set("경제")
        self.ai_subject.pack(side="left", padx=4)

        r1 = tk.Frame(f_settings)
        r1.pack(fill="x", padx=8, pady=3)
        tk.Label(r1, text="난이도:", width=10, anchor="w").pack(side="left")
        self.ai_difficulty = ttk.Combobox(r1, values=DIFFICULTY_LIST, width=40, state="readonly")
        self.ai_difficulty.set(DIFFICULTY_LIST[3])
        self.ai_difficulty.pack(side="left", padx=4)

        r2 = tk.Frame(f_settings)
        r2.pack(fill="x", padx=8, pady=3)
        tk.Label(r2, text="문항 수:", width=10, anchor="w").pack(side="left")
        self.ai_num_q = ttk.Combobox(r2, values=["5", "10", "15", "20", "25", "30"], width=8)
        self.ai_num_q.set("10")
        self.ai_num_q.pack(side="left", padx=4)

        r3 = tk.Frame(f_settings)
        r3.pack(fill="x", padx=8, pady=3)
        tk.Label(r3, text="범위/단원:", width=10, anchor="w").pack(side="left")
        self.ai_scope = tk.Entry(r3, width=50)
        self.ai_scope.pack(side="left", padx=4)
        tk.Label(r3, text="(예: 수요공급, 시장균형, 비교우위)", fg="gray", font=("맑은 고딕", 8)).pack(side="left")

        r4 = tk.Frame(f_settings)
        r4.pack(fill="x", padx=8, pady=3)
        tk.Label(r4, text="추가 지시:", width=10, anchor="w").pack(side="left")
        self.ai_extra = tk.Entry(r4, width=50)
        self.ai_extra.pack(side="left", padx=4)
        tk.Label(r4, text="(예: 그래프 문항 3개 포함)", fg="gray", font=("맑은 고딕", 8)).pack(side="left")

        # 버튼
        bf = tk.Frame(self.tab_ai)
        bf.pack(fill="x", padx=12, pady=6)
        tk.Button(bf, text="AI 명령어 생성", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=18, height=2,
                  command=self._generate_ai_prompt).pack(side="left", padx=6)
        tk.Button(bf, text="클립보드에 복사", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=18, height=2,
                  command=self._copy_ai_prompt).pack(side="left", padx=6)
        tk.Button(bf, text="AI 결과 -> 시험지 탭으로", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=22, height=2,
                  command=self._paste_ai_result).pack(side="left", padx=6)

        # 안내
        info_frame = tk.Frame(self.tab_ai, bg="#f0f9ff", relief="groove", bd=1)
        info_frame.pack(fill="x", padx=12, pady=(0, 4))
        tk.Label(info_frame, text="사용법: [AI 명령어 생성] -> [클립보드에 복사] -> ChatGPT/Claude에 붙여넣기 -> AI 출력 복사 -> [AI 결과 -> 시험지 탭으로]",
                 font=("맑은 고딕", 9), fg="#1e40af", bg="#f0f9ff", wraplength=900, justify="left").pack(padx=8, pady=4)

        # 프롬프트 미리보기
        tk.Label(self.tab_ai, text="생성된 AI 명령어 미리보기:", font=("맑은 고딕", 10, "bold"), fg="#7c3aed").pack(anchor="w", padx=12, pady=(4, 0))
        self.txt_ai_prompt = tk.Text(self.tab_ai, wrap=tk.WORD, height=20, font=("Consolas", 9), bg="#faf5ff")
        scroll_ai = tk.Scrollbar(self.tab_ai, command=self.txt_ai_prompt.yview)
        self.txt_ai_prompt.configure(yscrollcommand=scroll_ai.set)
        scroll_ai.pack(side="right", fill="y", padx=(0, 12))
        self.txt_ai_prompt.pack(fill="both", expand=True, padx=12, pady=4)

    def _generate_ai_prompt(self):
        subject = self.ai_subject.get()
        difficulty = self.ai_difficulty.get()
        num_q = self.ai_num_q.get()
        scope = self.ai_scope.get().strip() or "전 범위"
        extra = self.ai_extra.get().strip() or "없음"

        prompt = AI_MASTER_PROMPT_TEMPLATE.format(
            subject=subject,
            difficulty=difficulty,
            num_questions=num_q,
            scope=scope,
            extra=extra,
        )
        self.txt_ai_prompt.delete("1.0", tk.END)
        self.txt_ai_prompt.insert("1.0", prompt)
        self._set_status(f"AI 명령어 생성 완료 ({subject}, {num_q}문항)")

    def _copy_ai_prompt(self):
        prompt = self.txt_ai_prompt.get("1.0", tk.END).strip()
        if not prompt:
            self._generate_ai_prompt()
            prompt = self.txt_ai_prompt.get("1.0", tk.END).strip()
        if prompt:
            self.root.clipboard_clear()
            self.root.clipboard_append(prompt)
            self.root.update()
            messagebox.showinfo("복사 완료",
                                "AI 명령어가 클립보드에 복사되었습니다!\n\n"
                                "ChatGPT 또는 Claude에 붙여넣기(Ctrl+V)하세요.")

    def _paste_ai_result(self):
        """클립보드에서 AI 결과를 가져와 시험지/답안지 영역에 자동 분배."""
        try:
            clipboard = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showwarning("경고", "클립보드가 비어있습니다.\nAI 출력을 먼저 복사하세요.")
            return

        if not clipboard.strip():
            messagebox.showwarning("경고", "클립보드가 비어있습니다.")
            return

        # AI 출력에서 시험지/답안지 자동 분리
        # 패턴: "정답 및 해설", "교사용 정답", "[출력 2]" 등으로 분리
        split_patterns = [
            r"\n\s*\[출력\s*2\s*[—\-]\s*교사용",
            r"\n\s*#{1,3}\s*(?:정답\s*및\s*해설|교사용\s*정답)",
            r"\n\s*(?:정답\s*및\s*해설|교사용\s*정답\s*및\s*해설)\s*\n",
            r"\n\s*[-=]{5,}\s*\n\s*(?:정답|해설)",
        ]

        exam_part = clipboard
        ans_part = ""
        for pat in split_patterns:
            m = re.search(pat, clipboard, re.IGNORECASE)
            if m:
                exam_part = clipboard[:m.start()].strip()
                ans_part = clipboard[m.start():].strip()
                break

        # 시험지 영역에 삽입
        self.txt_exam.delete("1.0", tk.END)
        self.txt_exam.insert("1.0", exam_part)

        # 답안지 영역에 삽입
        if ans_part:
            self.txt_ans.delete("1.0", tk.END)
            self.txt_ans.insert("1.0", ans_part)

        self.tabs.select(self.tab_main)
        parsed = parse_exam_text(exam_part)
        self._set_status(f"AI 결과 붙여넣기 완료: {len(parsed)}문항 감지, 답안지 {'있음' if ans_part else '없음'}")

    # --- TAB 6: DB 도구 ---
    def _build_db_tab(self):
        # DB 폴더 설정
        lf_folder = tk.LabelFrame(self.tab_db, text=" DB 폴더 설정 ", font=("맑은 고딕", 10, "bold"))
        lf_folder.pack(fill="x", padx=12, pady=6)
        r0 = tk.Frame(lf_folder)
        r0.pack(fill="x", padx=8, pady=4)
        tk.Label(r0, text="DB 폴더:", width=10, anchor="w").pack(side="left")
        self.db_dir_var = tk.StringVar(value=self.cfg.get("db_root", os.path.join(SCRIPT_DIR, "db")))
        tk.Entry(r0, textvariable=self.db_dir_var, width=45).pack(side="left", padx=4)
        tk.Button(r0, text="찾아보기", command=self._browse_db_dir).pack(side="left", padx=2)

        # DB 작업 버튼
        lf_actions = tk.LabelFrame(self.tab_db, text=" DB 작업 ", font=("맑은 고딕", 10, "bold"))
        lf_actions.pack(fill="x", padx=12, pady=6)
        bf = tk.Frame(lf_actions)
        bf.pack(padx=8, pady=8)
        tk.Button(bf, text="DB 저장\n(파싱 결과 저장)", bg="#16a34a", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=16, height=3,
                  command=self._save_to_db).pack(side="left", padx=6)
        tk.Button(bf, text="문제은행\n(브라우저 열기)", bg="#2563eb", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=16, height=3,
                  command=self._open_qbank_window).pack(side="left", padx=6)
        tk.Button(bf, text="DB 백업\n(zip 생성)", bg="#7c3aed", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=16, height=3,
                  command=self._backup_db).pack(side="left", padx=6)
        tk.Button(bf, text="DB 복원\n(zip에서)", bg="#dc2626", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=16, height=3,
                  command=self._restore_db).pack(side="left", padx=6)
        tk.Button(bf, text="JSON Export\n(시험지 JSON)", bg="#0891b2", fg="white",
                  font=("맑은 고딕", 10, "bold"), width=16, height=3,
                  command=self._export_json).pack(side="left", padx=6)

        # 머리글/바닥글 설정
        lf_hf = tk.LabelFrame(self.tab_db, text=" 머리글 / 바닥글 설정 (DOCX 적용) ", font=("맑은 고딕", 10, "bold"))
        lf_hf.pack(fill="x", padx=12, pady=6)

        r1 = tk.Frame(lf_hf)
        r1.pack(fill="x", padx=8, pady=2)
        tk.Label(r1, text="학원명:", width=10, anchor="w").pack(side="left")
        self.hf_academy = tk.Entry(r1, width=30)
        self.hf_academy.pack(side="left", padx=4)
        self.hf_academy.insert(0, self.cfg.get("academy_name", ""))
        tk.Label(r1, text="시험명:", width=8, anchor="w").pack(side="left", padx=(8, 0))
        self.hf_exam = tk.Entry(r1, width=25)
        self.hf_exam.pack(side="left", padx=4)
        self.hf_exam.insert(0, self.cfg.get("exam_name", ""))

        r2 = tk.Frame(lf_hf)
        r2.pack(fill="x", padx=8, pady=2)
        tk.Label(r2, text="연락처:", width=10, anchor="w").pack(side="left")
        self.hf_contact = tk.Entry(r2, width=30)
        self.hf_contact.pack(side="left", padx=4)
        self.hf_contact.insert(0, self.cfg.get("contact_info", ""))
        tk.Label(r2, text="저작권:", width=8, anchor="w").pack(side="left", padx=(8, 0))
        self.hf_copyright = tk.Entry(r2, width=25)
        self.hf_copyright.pack(side="left", padx=4)
        self.hf_copyright.insert(0, self.cfg.get("copyright_text", ""))

        r3 = tk.Frame(lf_hf)
        r3.pack(fill="x", padx=8, pady=4)
        self.hf_page_var = tk.BooleanVar(value=self.cfg.get("show_page_number", False))
        tk.Checkbutton(r3, text="페이지 번호 표시 (Page X)", variable=self.hf_page_var).pack(side="left", padx=(76, 10))
        self.hf_date_var = tk.BooleanVar(value=self.cfg.get("show_date", False))
        tk.Checkbutton(r3, text="날짜 표시", variable=self.hf_date_var).pack(side="left", padx=10)
        tk.Button(r3, text="설정 저장", bg="#475569", fg="white",
                  font=("맑은 고딕", 9, "bold"), command=self._save_hf_config).pack(side="right", padx=8)

        # 미리보기 안내
        info = tk.Label(self.tab_db,
                        text="머리글 예시: 학원명 | 시험명 | 날짜 | Page X\n바닥글 예시: 저작권 | 연락처 | Page X\n(설정 저장 후 시험지 생성 시 자동 적용됩니다)",
                        font=("맑은 고딕", 9), fg="#64748b", bg="#f1f5f9", relief="groove", padx=8, pady=6, justify="left")
        info.pack(fill="x", padx=12, pady=(0, 6))

    # --- DB 작업 메서드 ---
    def _get_db_root(self):
        return self.db_dir_var.get() if hasattr(self, "db_dir_var") else self.cfg.get("db_root", os.path.join(SCRIPT_DIR, "db"))

    def _browse_db_dir(self):
        d = filedialog.askdirectory(title="DB 폴더 선택")
        if d:
            self.db_dir_var.set(d)
            self.cfg["db_root"] = d
            save_config(self.cfg)

    def _save_hf_config(self):
        self.cfg["academy_name"] = self.hf_academy.get().strip()
        self.cfg["exam_name"] = self.hf_exam.get().strip()
        self.cfg["contact_info"] = self.hf_contact.get().strip()
        self.cfg["copyright_text"] = self.hf_copyright.get().strip()
        self.cfg["show_page_number"] = self.hf_page_var.get()
        self.cfg["show_date"] = self.hf_date_var.get()
        self.cfg["db_root"] = self.db_dir_var.get()
        save_config(self.cfg)
        self._set_status("머리글/바닥글 설정 저장 완료")

    def _get_hf_config(self):
        return {
            "academy_name": self.cfg.get("academy_name", ""),
            "exam_name": self.cfg.get("exam_name", ""),
            "contact_info": self.cfg.get("contact_info", ""),
            "copyright_text": self.cfg.get("copyright_text", ""),
            "show_page_number": self.cfg.get("show_page_number", False),
            "show_date": self.cfg.get("show_date", False),
        }

    def _save_to_db(self):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        if not raw:
            return messagebox.showwarning("경고", "문제 텍스트를 먼저 입력하세요.")
        parsed = parse_exam_text(raw)
        if not parsed:
            return messagebox.showerror("오류", "파싱된 문제가 없습니다.")
        title = self.ent_title.get().strip() or "untitled"
        tag_str = simpledialog.askstring("태그", "저장 태그 (쉼표 구분):", initialvalue="", parent=self.root)
        if tag_str is None:
            return
        tags = [t.strip() for t in tag_str.split(",") if t.strip()]
        db_root = self._get_db_root()
        items = exam_db.parsed_to_db_items(parsed, source_title=title, tags=tags)
        for item in items:
            exam_db.save_item(db_root, item)
            exam_db.update_index_entry(db_root, item)
        q_total = sum(len(it.get("questions", [])) for it in items)
        messagebox.showinfo("DB 저장", f"{len(items)}개 아이템 ({q_total}문항) 저장 완료\nDB: {db_root}")
        self._set_status(f"DB 저장: {q_total}문항 → {db_root}")

    def _open_qbank_window(self):
        db_root = self._get_db_root()
        exam_db.ensure_db_dirs(db_root)

        def on_generate_from_db(items):
            exam_data = exam_db.db_items_to_exam_data(items)
            ans_text = exam_db.db_items_to_answer_text(items)
            self.txt_exam.delete("1.0", tk.END)
            # 재구성된 텍스트 형식으로 입력
            lines = []
            for q in exam_data:
                body = f"{q['num']}. {q['text']}"
                if q.get("jesi"):
                    body += f"\n<제시문>{q['jesi']}</제시문>"
                if q.get("graph_tag"):
                    body += f"\n{q['graph_tag']}"
                if q.get("bogi"):
                    body += f"\n<보기>\n{q['bogi']}"
                if q.get("choices"):
                    for j, ch in enumerate(q["choices"]):
                        if ch:
                            body += f"\n{CIRCLE_NUMS[j]} {ch}"
                lines.append(body)
            self.txt_exam.insert("1.0", "\n\n".join(lines))
            self.txt_ans.delete("1.0", tk.END)
            self.txt_ans.insert("1.0", ans_text)
            self.tabs.select(self.tab_main)
            self._set_status(f"DB에서 {len(exam_data)}문항 불러옴")

        QuestionBankWindow(self.root, db_root, on_generate_callback=on_generate_from_db)

    def _backup_db(self):
        db_root = self._get_db_root()
        try:
            path = exam_db.create_backup(db_root, CONFIG_FILE)
            messagebox.showinfo("백업 완료", f"백업 저장: {path}")
            self._set_status(f"DB 백업 완료: {path}")
        except Exception as e:
            messagebox.showerror("백업 오류", str(e))

    def _restore_db(self):
        db_root = self._get_db_root()
        zip_path = filedialog.askopenfilename(
            title="백업 zip 선택",
            filetypes=[("Zip", "*.zip")],
            initialdir=os.path.join(db_root, "backups"),
        )
        if not zip_path:
            return
        if not messagebox.askyesno("확인", f"복원하시겠습니까?\n기존 데이터는 안전 복사됩니다.\n\n{zip_path}"):
            return
        try:
            safety = exam_db.restore_backup(db_root, zip_path)
            exam_db.build_index(db_root)
            messagebox.showinfo("복원 완료", f"복원 완료!\n안전 복사: {safety}")
            self._set_status("DB 복원 완료")
        except Exception as e:
            messagebox.showerror("복원 오류", str(e))

    def _export_json(self):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        if not raw:
            return messagebox.showwarning("경고", "문제 텍스트를 먼저 입력하세요.")
        parsed = parse_exam_text(raw)
        if not parsed:
            return messagebox.showerror("오류", "파싱된 문제가 없습니다.")
        target_dir = self.dir_var.get()
        title = self.ent_title.get().strip() or "exam"
        today = datetime.datetime.now().strftime("%Y%m%d")
        filename = f"{title}_{today}"
        try:
            path = exam_db.export_exam_json(parsed, target_dir, filename,
                                             display_title=title)
            messagebox.showinfo("Export 완료", f"저장: {path}")
            self._set_status(f"JSON Export: {path}")
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # --- 핵심: 생성 로직 ---
    def _get_target_dir(self, student_name=""):
        base = self.dir_var.get()
        title = self.ent_title.get().strip()
        name = student_name
        if not name and "_" in title:
            name = title.split("_")[0].strip()
        if name:
            return os.path.join(base, name)
        return base

    def _generate(self, fmt):
        raw_exam = self.txt_exam.get("1.0", tk.END).strip()
        raw_ans = self.txt_ans.get("1.0", tk.END).strip()
        if not raw_exam:
            return messagebox.showwarning("경고", "문제를 입력하세요!")
        parsed = parse_exam_text(raw_exam)
        if not parsed:
            return messagebox.showerror("오류", "파싱 가능한 문제가 없습니다. 형식을 확인하세요.")
        title_raw = self.ent_title.get().strip()
        display_title = title_raw if title_raw else "사탐 모의고사"
        today = datetime.datetime.now().strftime("%Y%m%d")
        filename_base = f"{title_raw}_{today}" if title_raw else f"교재_{today}"
        font_n = self.font_combo.get()
        font_s = self.size_combo.get()
        logo = self.logo_var.get()
        watermark = self.ent_watermark.get().strip()
        shuffle = self.var_shuffle.get()

        batch_raw = self.ent_batch.get().strip()
        students = [s.strip() for s in batch_raw.split(",") if s.strip()] if batch_raw else [""]
        self._save_cfg()

        generated_paths = []
        try:
            for student in students:
                target_dir = self._get_target_dir(student)
                os.makedirs(target_dir, exist_ok=True)
                data = list(parsed)
                if shuffle:
                    data = data[:]
                    random.shuffle(data)
                fn = f"{student}_{filename_base}" if student else filename_base
                header_txt = f"{student} | {display_title}" if student else display_title
                footer_txt = datetime.datetime.now().strftime("%Y-%m-%d")
                hf_cfg = self._get_hf_config() if hasattr(self, "_get_hf_config") else None
                exam_dx, exam_px = create_exam_docx(
                    target_dir, fn, data, font_n, font_s, logo,
                    display_title, watermark=watermark,
                    header=header_txt, footer=footer_txt,
                    hf_config=hf_cfg,
                )
                ans_dx, ans_px = None, None
                if raw_ans:
                    ans_dx, ans_px = create_answer_docx(
                        target_dir, fn, raw_ans, data, font_n, font_s,
                        logo, display_title, hf_config=hf_cfg,
                    )
                if fmt in ("pdf", "both"):
                    pairs = [(exam_dx, exam_px)]
                    if ans_dx:
                        pairs.append((ans_dx, ans_px))
                    convert_docx_pairs_to_pdf(pairs)
                    if fmt == "pdf":
                        for d in (exam_dx, ans_dx):
                            if d and os.path.exists(d):
                                os.remove(d)
                generated_paths.append(target_dir)
                _add_history({
                    "title": display_title,
                    "student": student or "(단일)",
                    "q_count": len(data),
                    "format": fmt,
                    "path": target_dir,
                })

            summary = f"총 {len(students)}명 x {len(parsed)}문항 생성 완료!" if len(students) > 1 else f"{len(parsed)}문항 생성 완료!"
            messagebox.showinfo("완료", f"{summary}\n\n저장: {generated_paths[0]}")
            _clear_draft()
            self._refresh_history()
            self._update_status()
            if generated_paths:
                self._open_dir(generated_paths[0])
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # --- 유틸리티 ---
    def _browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.dir_var.set(d)
            self._save_cfg()

    def _browse_logo(self):
        f = filedialog.askopenfilename(filetypes=[("이미지", "*.png *.jpg *.jpeg *.bmp")])
        if f:
            self.logo_var.set(f)
            self._save_cfg()

    def _open_dir(self, path):
        import subprocess as _sp
        if not os.path.exists(path):
            path = self.dir_var.get()
        if os.path.exists(path):
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                _sp.Popen(["open", path])
            else:
                _sp.Popen(["xdg-open", path])

    def _open_target_dir(self):
        self._open_dir(self._get_target_dir())

    def _save_cfg(self):
        self.cfg["last_dir"] = self.dir_var.get()
        self.cfg["last_logo"] = self.logo_var.get()
        self.cfg["font_name"] = self.font_combo.get()
        self.cfg["font_size"] = self.size_combo.get()
        save_config(self.cfg)

    def _set_status(self, text):
        self.lbl_status.config(text=text)

    def _update_status(self):
        raw = self.txt_exam.get("1.0", tk.END).strip()
        q_count = len(parse_exam_text(raw)) if raw else 0
        char_count = len(raw)
        bank_count = len(load_qbank())
        hist_count = len(_load_history())
        self._set_status(
            f"현재 입력: {q_count}문항 ({char_count}자)  |  "
            f"문제 은행: {bank_count}문항  |  "
            f"히스토리: {hist_count}건"
        )

    def _on_text_change(self, event=None):
        self._update_status()

    def _save_draft_now(self, event=None):
        _save_draft(
            self.txt_exam.get("1.0", tk.END).strip(),
            self.txt_ans.get("1.0", tk.END).strip(),
            self.ent_title.get().strip(),
        )
        self._set_status("임시저장 완료")

    def _start_autosave(self):
        def _tick():
            exam = self.txt_exam.get("1.0", tk.END).strip()
            if exam:
                _save_draft(exam, self.txt_ans.get("1.0", tk.END).strip(), self.ent_title.get().strip())
            self.root.after(30000, _tick)
        self.root.after(30000, _tick)

    def _restore_draft(self):
        draft = _load_draft()
        if draft and draft.get("exam"):
            saved_at = draft.get("saved_at", "")[:19].replace("T", " ")
            if messagebox.askyesno("임시저장 복원", f"이전 임시저장({saved_at})이 있습니다.\n복원하시겠습니까?"):
                self.txt_exam.insert("1.0", draft["exam"])
                if draft.get("answer"):
                    self.txt_ans.insert("1.0", draft["answer"])
                if draft.get("title"):
                    self.ent_title.insert(0, draft["title"])
                self._update_status()

    def _on_file_drop(self, event):
        """Drag & Drop으로 파일을 받아 텍스트 입력 영역에 삽입."""
        files = self.root.tk.splitlist(event.data)
        if not files:
            return
        file_path = files[0]
        ext = os.path.splitext(file_path)[1].lower()
        extracted_text = ""
        try:
            if ext == ".txt":
                for enc in ("utf-8", "cp949", "euc-kr"):
                    try:
                        with open(file_path, "r", encoding=enc) as f:
                            extracted_text = f.read()
                        break
                    except UnicodeDecodeError:
                        continue
            elif ext == ".docx":
                doc = docx.Document(file_path)
                extracted_text = "\n".join(p.text for p in doc.paragraphs)
            else:
                messagebox.showwarning("DnD", f"지원하지 않는 파일 형식: {ext}\n(.txt, .docx만 지원)")
                return
            if extracted_text:
                current = self.txt_exam.get("1.0", tk.END).strip()
                if current:
                    self.txt_exam.insert(tk.END, "\n\n" + extracted_text)
                else:
                    self.txt_exam.insert("1.0", extracted_text)
                self._update_status()
                messagebox.showinfo("DnD", f"파일 로드 완료: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("DnD 오류", f"파일 파싱 중 에러: {e}")

    def _on_close(self):
        exam = self.txt_exam.get("1.0", tk.END).strip()
        if exam:
            _save_draft(exam, self.txt_ans.get("1.0", tk.END).strip(), self.ent_title.get().strip())
        self.root.destroy()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 엔트리포인트
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
if __name__ == "__main__":
    if TkinterDnD:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
        print("Info: tkinterdnd2 미설치 — Drag & Drop 비활성화")
    app = SutamMakerApp(root)
    root.mainloop()
