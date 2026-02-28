"""
학생 모델 — CRUD, 출제 이력, 오답 통계
"""

import datetime
from exam_bank.models.database import db_conn


# ── 학생 CRUD ─────────────────────────────────

def create_student(cfg, name, grade="", target="", level=""):
    """학생 추가. 반환: student_id."""
    now = datetime.datetime.now().isoformat()
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO students (name, grade, target, level, created_at) VALUES (?,?,?,?,?)",
            (name, grade, target, level, now)
        )
        return cur.lastrowid


def update_student(cfg, sid, name=None, grade=None, target=None, level=None):
    """학생 정보 수정."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        sets, params = [], []
        if name is not None:
            sets.append("name=?"); params.append(name)
        if grade is not None:
            sets.append("grade=?"); params.append(grade)
        if target is not None:
            sets.append("target=?"); params.append(target)
        if level is not None:
            sets.append("level=?"); params.append(level)
        if sets:
            params.append(sid)
            cur.execute(f"UPDATE students SET {','.join(sets)} WHERE id=?", params)


def delete_student(cfg, sid):
    """학생 삭제 (CASCADE: 시험기록, 분석기록도 삭제)."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM exam_records WHERE student_id=?", (sid,))
        cur.execute("DELETE FROM analysis_records WHERE student_id=?", (sid,))
        cur.execute("DELETE FROM students WHERE id=?", (sid,))


def list_students(cfg):
    """전체 학생 목록."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        rows = cur.execute("SELECT * FROM students ORDER BY id").fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_student(cfg, sid):
    """학생 1명 조회."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        row = cur.execute("SELECT * FROM students WHERE id=?", (sid,)).fetchone()
        if not row:
            return None
        cols = [d[0] for d in cur.description]
        return dict(zip(cols, row))


# ── 시험 출제 기록 ────────────────────────────

def create_exam_record(cfg, student_id, question_ids, exam_title="", shuffle_order=None):
    """시험 출제 기록 생성. question_ids: 출제할 문항 ID 리스트.
    shuffle_order: 문항 순서 (None이면 원래 순서).
    반환: exam_id."""
    now = datetime.datetime.now()
    order = shuffle_order if shuffle_order else question_ids

    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO exam_records
               (student_id, exam_date, exam_title, total_questions, created_at)
               VALUES (?,?,?,?,?)""",
            (student_id, now.strftime("%Y-%m-%d"), exam_title, len(order), now.isoformat())
        )
        exam_id = cur.lastrowid

        for idx, qid in enumerate(order, 1):
            cur.execute(
                "INSERT INTO exam_items (exam_id, question_id, display_order) VALUES (?,?,?)",
                (exam_id, qid, idx)
            )
            # usage_count 증가
            cur.execute("UPDATE questions SET usage_count = usage_count + 1 WHERE id=?", (qid,))

        return exam_id


def get_exam_record(cfg, exam_id):
    """시험 기록 상세 조회."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        row = cur.execute("SELECT * FROM exam_records WHERE id=?", (exam_id,)).fetchone()
        if not row:
            return None
        cols = [d[0] for d in cur.description]
        record = dict(zip(cols, row))

        items = cur.execute(
            """SELECT ei.*, q.content as q_content, q.q_num, q.answer as correct_answer,
                      q.choices, q.choice_type, q.difficulty,
                      p.content as passage_content, p.extra_tags
               FROM exam_items ei
               JOIN questions q ON ei.question_id = q.id
               JOIN passages p ON q.passage_id = p.id
               WHERE ei.exam_id=?
               ORDER BY ei.display_order""",
            (exam_id,)
        ).fetchall()
        item_cols = [d[0] for d in cur.description]
        record["items"] = [dict(zip(item_cols, i)) for i in items]
        return record


def list_exam_records(cfg, student_id=None):
    """시험 기록 목록."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        if student_id:
            rows = cur.execute(
                "SELECT * FROM exam_records WHERE student_id=? ORDER BY id DESC",
                (student_id,)
            ).fetchall()
        else:
            rows = cur.execute("SELECT * FROM exam_records ORDER BY id DESC").fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


# ── 답안 기입 & 채점 ─────────────────────────

def submit_answers(cfg, exam_id, answers):
    """학생 답안 기입 + 자동 채점.
    answers: {question_id: "학생답안", ...}
    반환: {"total": N, "correct": N, "wrong": N, "score": float}
    """
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        total = 0
        correct = 0

        items = cur.execute(
            """SELECT ei.id, ei.question_id, q.answer as correct_answer
               FROM exam_items ei
               JOIN questions q ON ei.question_id = q.id
               WHERE ei.exam_id=?""",
            (exam_id,)
        ).fetchall()

        for ei_id, qid, correct_ans in items:
            student_ans = answers.get(qid, answers.get(str(qid), ""))
            if not student_ans:
                continue

            total += 1
            is_correct = 1 if student_ans.strip() == (correct_ans or "").strip() else 0
            if is_correct:
                correct += 1

            cur.execute(
                "UPDATE exam_items SET student_answer=?, is_correct=? WHERE id=?",
                (student_ans, is_correct, ei_id)
            )

        score = round((correct / total * 100), 1) if total > 0 else 0
        cur.execute(
            "UPDATE exam_records SET is_graded=1, score=? WHERE id=?",
            (score, exam_id)
        )

    return {"total": total, "correct": correct, "wrong": total - correct, "score": score}


# ── 출제 이력 조회 ────────────────────────────

def get_question_history(cfg, question_id, student_id=None):
    """특정 문항의 출제 이력. 반환: [{student_name, exam_date, is_correct, ...}, ...]"""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        query = """SELECT s.name as student_name, er.exam_date, ei.is_correct,
                          ei.student_answer, er.id as exam_id
                   FROM exam_items ei
                   JOIN exam_records er ON ei.exam_id = er.id
                   JOIN students s ON er.student_id = s.id
                   WHERE ei.question_id=?"""
        params = [question_id]
        if student_id:
            query += " AND er.student_id=?"
            params.append(student_id)
        query += " ORDER BY er.exam_date DESC"
        rows = cur.execute(query, params).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_student_question_stats(cfg, student_id, question_id):
    """학생의 특정 문항 출제 통계. 반환: {count, correct, wrong, last_date}."""
    history = get_question_history(cfg, question_id, student_id)
    if not history:
        return {"count": 0, "correct": 0, "wrong": 0, "last_date": None}
    return {
        "count": len(history),
        "correct": sum(1 for h in history if h["is_correct"] == 1),
        "wrong": sum(1 for h in history if h["is_correct"] == 0),
        "last_date": history[0]["exam_date"] if history else None,
    }


# ── 오답 데이터 수집 ─────────────────────────

def get_student_wrong_answers(cfg, student_id, limit=50):
    """학생의 오답 문항 전체 조회. 최근순."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        rows = cur.execute(
            """SELECT ei.question_id, ei.student_answer, q.answer as correct_answer,
                      q.content as q_content, q.q_num, q.difficulty,
                      p.content as passage_content, p.extra_tags, p.category1,
                      p.school_year, er.exam_date
               FROM exam_items ei
               JOIN exam_records er ON ei.exam_id = er.id
               JOIN questions q ON ei.question_id = q.id
               JOIN passages p ON q.passage_id = p.id
               WHERE er.student_id=? AND ei.is_correct=0
               ORDER BY er.exam_date DESC
               LIMIT ?""",
            (student_id, limit)
        ).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_student_weakness_summary(cfg, student_id):
    """학생의 태그별 오답률 분석.
    반환: [{tag, total, wrong, wrong_rate}, ...] 오답률 높은 순."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        rows = cur.execute(
            """SELECT p.extra_tags, ei.is_correct
               FROM exam_items ei
               JOIN exam_records er ON ei.exam_id = er.id
               JOIN questions q ON ei.question_id = q.id
               JOIN passages p ON q.passage_id = p.id
               WHERE er.student_id=? AND ei.is_correct >= 0""",
            (student_id,)
        ).fetchall()

    # 태그별 집계
    tag_stats = {}
    for tags_str, is_correct in rows:
        tags = [t.strip() for t in (tags_str or "").split(",") if t.strip()]
        if not tags:
            tags = ["미분류"]
        for tag in tags:
            if tag not in tag_stats:
                tag_stats[tag] = {"total": 0, "wrong": 0}
            tag_stats[tag]["total"] += 1
            if is_correct == 0:
                tag_stats[tag]["wrong"] += 1

    result = []
    for tag, stat in tag_stats.items():
        rate = round(stat["wrong"] / stat["total"] * 100, 1) if stat["total"] > 0 else 0
        result.append({
            "tag": tag,
            "total": stat["total"],
            "wrong": stat["wrong"],
            "wrong_rate": rate,
        })
    result.sort(key=lambda x: x["wrong_rate"], reverse=True)
    return result
