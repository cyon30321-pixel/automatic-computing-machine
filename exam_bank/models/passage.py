"""
지문/문항 모델 — CRUD, 검색, 중복 감지
"""

import json
import datetime
from exam_bank.models.database import db_conn
from exam_bank.models.validators import compute_content_hash, normalize_choices, detect_choice_type


# ── 지문 CRUD ─────────────────────────────────

def create_passage(cfg, content, category1="", category2="", school_year="",
                   exam_year="", exam_month="", publisher="", extra_tags="",
                   answer_text="", questions=None):
    """지문 + 문항 일괄 저장. 반환: passage_id."""
    now = datetime.datetime.now().isoformat()
    q_texts = " ".join(q.get("content", "") for q in (questions or []))
    content_hash = compute_content_hash(content, q_texts)

    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO passages
               (content, category1, category2, school_year, exam_year,
                exam_month, publisher, extra_tags, answer_text,
                content_hash, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
            (content, category1, category2, school_year, exam_year,
             exam_month, publisher, extra_tags, answer_text,
             content_hash, now, now)
        )
        pid = cur.lastrowid

        for q in (questions or []):
            choices = json.dumps(normalize_choices(q.get("choices", [])), ensure_ascii=False)
            c_type = detect_choice_type(q.get("choices", []))
            cur.execute(
                """INSERT INTO questions
                   (passage_id, q_num, content, choices, choice_type,
                    answer, explanation, difficulty)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (pid, q.get("q_num", "-"), q.get("content", ""),
                 choices, c_type,
                 q.get("answer", ""), q.get("explanation", ""),
                 q.get("difficulty", 3))
            )
    return pid


def update_passage(cfg, pid, content=None, answer_text=None, questions=None, **kwargs):
    """지문 수정. questions가 주어지면 기존 문항 삭제 후 재등록."""
    now = datetime.datetime.now().isoformat()
    with db_conn(cfg) as conn:
        cur = conn.cursor()

        sets, params = ["updated_at=?"], [now]
        if content is not None:
            sets.append("content=?")
            params.append(content)
        if answer_text is not None:
            sets.append("answer_text=?")
            params.append(answer_text)
        for col in ("category1", "category2", "school_year", "exam_year",
                     "exam_month", "publisher", "extra_tags"):
            if col in kwargs:
                sets.append(f"{col}=?")
                params.append(kwargs[col])

        params.append(pid)
        cur.execute(f"UPDATE passages SET {','.join(sets)} WHERE id=?", params)

        if questions is not None:
            cur.execute("DELETE FROM questions WHERE passage_id=?", (pid,))
            for q in questions:
                choices = json.dumps(normalize_choices(q.get("choices", [])), ensure_ascii=False)
                c_type = detect_choice_type(q.get("choices", []))
                cur.execute(
                    """INSERT INTO questions
                       (passage_id, q_num, content, choices, choice_type,
                        answer, explanation, difficulty)
                       VALUES (?,?,?,?,?,?,?,?)""",
                    (pid, q.get("q_num", "-"), q.get("content", ""),
                     choices, c_type,
                     q.get("answer", ""), q.get("explanation", ""),
                     q.get("difficulty", 3))
                )


def delete_passage(cfg, pid):
    """지문 삭제 (CASCADE로 문항도 삭제)."""
    with db_conn(cfg) as conn:
        conn.cursor().execute("DELETE FROM passages WHERE id=?", (pid,))


def delete_question(cfg, qid):
    """개별 문항 삭제."""
    with db_conn(cfg) as conn:
        conn.cursor().execute("DELETE FROM questions WHERE id=?", (qid,))


def get_passage(cfg, pid):
    """지문 1개 조회. dict 반환."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        row = cur.execute("SELECT * FROM passages WHERE id=?", (pid,)).fetchone()
        if not row:
            return None
        cols = [d[0] for d in cur.description]
        return dict(zip(cols, row))


def get_questions_for_passage(cfg, pid):
    """지문에 속한 문항 리스트."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        rows = cur.execute(
            "SELECT * FROM questions WHERE passage_id=? ORDER BY id", (pid,)
        ).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_question(cfg, qid):
    """문항 1개 조회."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        row = cur.execute("SELECT * FROM questions WHERE id=?", (qid,)).fetchone()
        if not row:
            return None
        cols = [d[0] for d in cur.description]
        return dict(zip(cols, row))


def update_question_answer(cfg, qid, answer="", explanation=""):
    """문항 정답/해설 업데이트."""
    with db_conn(cfg) as conn:
        conn.cursor().execute(
            "UPDATE questions SET answer=?, explanation=? WHERE id=?",
            (answer, explanation, qid)
        )


def update_question_difficulty(cfg, qid, difficulty):
    """문항 난이도 업데이트."""
    with db_conn(cfg) as conn:
        conn.cursor().execute(
            "UPDATE questions SET difficulty=? WHERE id=?", (difficulty, qid)
        )


# ── 검색 ──────────────────────────────────────

def search_passages(cfg, search="", category1="", category2="", exam_year=""):
    """필터 기반 지문 검색."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        query = """SELECT id, category1, category2, school_year, exam_year,
                          exam_month, publisher, extra_tags, content, created_at, answer_text
                   FROM passages WHERE 1=1"""
        params = []

        if search:
            query += " AND (content LIKE ? OR extra_tags LIKE ? OR publisher LIKE ?)"
            params += [f"%{search}%"] * 3
        if category1 and category1 != "전체":
            query += " AND category1=?"
            params.append(category1)
        if category2 and category2 != "전체":
            query += " AND category2=?"
            params.append(category2)
        if exam_year and exam_year != "전체":
            query += " AND exam_year=?"
            params.append(exam_year)

        query += " ORDER BY id DESC"
        rows = cur.execute(query, params).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def find_duplicate(cfg, content):
    """내용 해시 기반 중복 검사. 중복이면 passage_id, 없으면 None."""
    h = compute_content_hash(content)
    with db_conn(cfg) as conn:
        row = conn.cursor().execute(
            "SELECT id FROM passages WHERE content_hash=?", (h,)
        ).fetchone()
        return row[0] if row else None


# ── 난이도 기반 자동 출제 ─────────────────────

def auto_select_questions(cfg, count=15, category1="", category2="",
                          school_year="", min_diff=1, max_diff=5):
    """조건에 맞는 문항 자동 추출. 사용횟수 적은 것 우선."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        query = """SELECT q.id, q.passage_id, q.q_num, q.content, q.difficulty, q.usage_count
                   FROM questions q
                   JOIN passages p ON q.passage_id = p.id
                   WHERE q.difficulty BETWEEN ? AND ?"""
        params = [min_diff, max_diff]

        if category1 and category1 != "전체":
            query += " AND p.category1=?"
            params.append(category1)
        if category2 and category2 != "전체":
            query += " AND p.category2=?"
            params.append(category2)
        if school_year and school_year != "전체":
            query += " AND p.school_year=?"
            params.append(school_year)

        query += " ORDER BY q.usage_count ASC, RANDOM()"
        query += " LIMIT ?"
        params.append(count)

        rows = cur.execute(query, params).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]
