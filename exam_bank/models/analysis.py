"""
분석 기록 모델 — 수동/자동 분석 CRUD
"""

import datetime
from exam_bank.models.database import db_conn


def save_analysis(cfg, student_id, weakness_tag, score, feedback="", source="manual"):
    """분석 기록 저장."""
    with db_conn(cfg) as conn:
        conn.cursor().execute(
            """INSERT INTO analysis_records
               (student_id, record_date, weakness_tag, score, feedback, source)
               VALUES (?,?,?,?,?,?)""",
            (student_id, datetime.date.today().isoformat(), weakness_tag, score, feedback, source)
        )


def save_auto_analysis(cfg, student_id):
    """오답 데이터 기반 자동 분석 저장."""
    from exam_bank.models.student import get_student_weakness_summary

    summary = get_student_weakness_summary(cfg, student_id)
    if not summary:
        return None

    # 상위 3개 취약점
    top = summary[:3]
    tags = ", ".join(t["tag"] for t in top)
    feedback_lines = []
    for t in top:
        feedback_lines.append(f"[{t['tag']}] 출제 {t['total']}회 / 오답 {t['wrong']}회 (오답률 {t['wrong_rate']}%)")

    avg_rate = round(sum(t["wrong_rate"] for t in top) / len(top), 1) if top else 0
    score = max(0, round(100 - avg_rate))

    feedback = "▶ 자동 분석 결과\n" + "\n".join(feedback_lines)

    save_analysis(cfg, student_id, tags, score, feedback, source="auto")
    return {"tags": tags, "score": score, "details": top}


def get_analysis_history(cfg, student_id):
    """학생 분석 기록 전체."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        rows = cur.execute(
            """SELECT * FROM analysis_records
               WHERE student_id=? ORDER BY id DESC""",
            (student_id,)
        ).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_latest_weakness(cfg, student_id):
    """최신 취약점 태그."""
    with db_conn(cfg) as conn:
        row = conn.cursor().execute(
            "SELECT weakness_tag FROM analysis_records WHERE student_id=? ORDER BY id DESC LIMIT 1",
            (student_id,)
        ).fetchone()
        return row[0] if row else None
