"""
분석 서비스 — 차트, CSV, AI 프롬프트 생성
"""

import os
import csv
import datetime
from collections import Counter

import matplotlib.pyplot as plt

from exam_bank.models.database import db_conn
from exam_bank.models.student import (
    get_student, get_student_wrong_answers, get_student_weakness_summary
)
from exam_bank.models.analysis import get_analysis_history


# ── 차트 ──────────────────────────────────────

def chart_progress(cfg, student_id, student_name=""):
    """성취도 추이 Line 차트."""
    history = get_analysis_history(cfg, student_id)
    if not history:
        return False, "데이터가 없습니다."

    history = list(reversed(history))  # 오래된 것부터
    dates = [f"{h['record_date'][5:]}({i+1}차)" for i, h in enumerate(history)]
    scores = [h["score"] for h in history]

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(dates, scores, marker="o", color="#2563eb", linewidth=2.5, markersize=8)
    ax.fill_between(dates, scores, alpha=0.12, color="#2563eb")
    ax.axhline(y=80, color="#dc2626", linestyle="--", alpha=0.6, label="목표 80점")
    for i, s in enumerate(scores):
        ax.annotate(f"{s}", (dates[i], s), textcoords="offset points",
                    xytext=(0, 10), ha="center", fontsize=9, fontweight="bold")
    ax.set_title(f"[{student_name}] 성취도 변화", fontsize=15, fontweight="bold", pad=15)
    ax.set_xlabel("평가 시기", fontsize=11)
    ax.set_ylabel("이해도 (100점)", fontsize=11)
    ax.set_ylim(0, 105)
    ax.legend()
    fig.tight_layout()
    plt.show()
    return True, ""


def chart_weakness(cfg, student_id, student_name=""):
    """취약점 분석 Bar 차트 (오답 데이터 기반)."""
    summary = get_student_weakness_summary(cfg, student_id)
    if not summary:
        # 레거시: analysis_records에서 가져오기
        history = get_analysis_history(cfg, student_id)
        if not history:
            return False, "데이터가 없습니다."
        counts = Counter(h["weakness_tag"] for h in history if h.get("weakness_tag"))
        items = counts.most_common()
        labels = [x[0] for x in items][::-1]
        values = [x[1] for x in items][::-1]
    else:
        # 오답 데이터 기반
        items = [(s["tag"], s["wrong_rate"]) for s in summary if s["total"] >= 1]
        labels = [x[0] for x in items][::-1]
        values = [x[1] for x in items][::-1]

    if not labels:
        return False, "데이터가 없습니다."

    fig, ax = plt.subplots(figsize=(9, 5))
    bars = ax.barh(labels, values, color="#7c3aed", alpha=0.85, height=0.6)
    for b in bars:
        label_text = f"{b.get_width():.0f}%" if max(values) > 10 else f"{int(b.get_width())}건"
        ax.text(b.get_width() + 0.5, b.get_y() + b.get_height() / 2,
                label_text, va="center", fontsize=11, fontweight="bold")
    ax.set_title(f"[{student_name}] 취약점 분석", fontsize=15, fontweight="bold", pad=15)
    ax.set_xlabel("오답률 (%)" if max(values) > 10 else "횟수", fontsize=11)
    fig.tight_layout()
    plt.show()
    return True, ""


# ── CSV Export ────────────────────────────────

def export_csv(cfg, student_id, student_name, target_dir):
    """학습 리포트 CSV 내보내기."""
    path = os.path.join(target_dir, f"[{student_name}]_학습리포트.csv")
    history = get_analysis_history(cfg, student_id)

    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["날짜", "학생", "취약점", "이해도", "소스", "분석 내용"])
        for r in history:
            w.writerow([r["record_date"], student_name, r["weakness_tag"],
                        f"{r['score']}점", r.get("source", "manual"),
                        str(r["feedback"]).strip()])
    return path


# ── AI 프롬프트 생성 ─────────────────────────

def generate_ai_prompt(cfg, student_id):
    """학생의 오답 데이터 기반 AI 맞춤 프롬프트 생성."""
    student = get_student(cfg, student_id)
    if not student:
        return None

    weakness = get_student_weakness_summary(cfg, student_id)
    wrong_answers = get_student_wrong_answers(cfg, student_id, limit=10)
    today = datetime.date.today().isoformat()

    # 프롬프트 구성
    lines = []
    lines.append("=" * 60)
    lines.append("AI 맞춤 문제 생성 프롬프트")
    lines.append("=" * 60)
    lines.append("")

    # 학생 프로필
    lines.append("[학생 프로필]")
    lines.append(f"이름: {student['name']}")
    if student.get("grade"):
        lines.append(f"학년: {student['grade']}")
    if student.get("target"):
        lines.append(f"목표: {student['target']}")
    lines.append(f"분석 기준일: {today}")
    lines.append("")

    # 취약점 분석
    if weakness:
        lines.append("[누적 오답 분석]")
        for w in weakness[:5]:
            status = ""
            if w["wrong_rate"] >= 70:
                status = " ← 최약점"
            elif w["wrong_rate"] <= 30:
                status = " ← 개선됨"
            lines.append(
                f"- {w['tag']}: {w['total']}회 출제 / {w['wrong']}회 오답 "
                f"(오답률 {w['wrong_rate']}%){status}"
            )
        lines.append("")

    # 최근 오답 문항
    if wrong_answers:
        lines.append("[최근 틀린 문항 원문]")
        for i, wa in enumerate(wrong_answers[:5], 1):
            lines.append(
                f"{i}. ({wa['exam_date']}) {wa['q_content'][:80]}"
            )
            lines.append(
                f"   → 학생 답: {wa['student_answer']} / 정답: {wa['correct_answer']}"
            )
            if wa.get("extra_tags"):
                lines.append(f"   태그: {wa['extra_tags']}")
            lines.append("")

    # 요청
    top_tags = ", ".join(w["tag"] for w in (weakness or [])[:3] if w["wrong_rate"] >= 40)
    grade = student.get("grade", "")
    level_hint = f"{grade} 수준 " if grade else ""

    lines.append("[요청]")
    if top_tags:
        lines.append(
            f"위 학생의 취약점({top_tags})을 집중 훈련할 수 있는 "
            f"{level_hint}영어 문제 10문항을 만들어주세요."
        )
    else:
        lines.append(
            f"위 학생의 오답 패턴을 분석하여 "
            f"{level_hint}보충 문제 10문항을 만들어주세요."
        )
    lines.append("")
    lines.append("각 문항에는 다음을 포함해주세요:")
    lines.append("1. 지문 (필요 시)")
    lines.append("2. 문제 + 5지선다 선지")
    lines.append("3. 정답 및 해설")
    lines.append("4. 난이도 (1~5)")
    lines.append("")
    lines.append("=" * 60)

    return "\n".join(lines)
