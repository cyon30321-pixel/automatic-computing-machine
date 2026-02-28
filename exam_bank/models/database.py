"""
DB 코어 — 연결 관리, 마이그레이션, 컨텍스트 매니저
"""

import os
import sqlite3
from contextlib import contextmanager

from exam_bank.config import get_db_path


@contextmanager
def db_conn(cfg):
    conn = sqlite3.connect(get_db_path(cfg))
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
    except Exception:
        conn.rollback()
        raise
    else:
        conn.commit()
    finally:
        conn.close()


def init_db(cfg):
    schema_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), "resources", "schema.sql")
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        if os.path.exists(schema_path):
            with open(schema_path, "r", encoding="utf-8") as f:
                cur.executescript(f.read())
        else:
            _create_tables_inline(cur)
        _migrate(cur)


def _create_tables_inline(cur):
    cur.execute("""CREATE TABLE IF NOT EXISTS passages (
        id INTEGER PRIMARY KEY AUTOINCREMENT, content TEXT NOT NULL,
        category1 TEXT DEFAULT '', category2 TEXT DEFAULT '',
        school_year TEXT DEFAULT '', exam_year TEXT DEFAULT '',
        exam_month TEXT DEFAULT '', publisher TEXT DEFAULT '',
        extra_tags TEXT DEFAULT '', answer_text TEXT DEFAULT '',
        content_hash TEXT DEFAULT '', created_at TEXT DEFAULT '', updated_at TEXT DEFAULT '')""")
    cur.execute("""CREATE TABLE IF NOT EXISTS questions (
        id INTEGER PRIMARY KEY AUTOINCREMENT, passage_id INTEGER NOT NULL,
        q_num TEXT DEFAULT '-', content TEXT DEFAULT '', choices TEXT DEFAULT '[]',
        choice_type TEXT DEFAULT 'num', answer TEXT DEFAULT '',
        explanation TEXT DEFAULT '', difficulty INTEGER DEFAULT 3,
        usage_count INTEGER DEFAULT 0,
        FOREIGN KEY (passage_id) REFERENCES passages(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL,
        target TEXT DEFAULT '', level TEXT DEFAULT '', grade TEXT DEFAULT '',
        created_at TEXT DEFAULT '')""")
    cur.execute("""CREATE TABLE IF NOT EXISTS exam_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER NOT NULL,
        exam_date TEXT NOT NULL, exam_title TEXT DEFAULT '',
        total_questions INTEGER DEFAULT 0, is_graded INTEGER DEFAULT 0,
        score REAL DEFAULT 0, created_at TEXT DEFAULT '',
        FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS exam_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT, exam_id INTEGER NOT NULL,
        question_id INTEGER NOT NULL, display_order INTEGER DEFAULT 0,
        student_answer TEXT DEFAULT '', is_correct INTEGER DEFAULT -1,
        FOREIGN KEY (exam_id) REFERENCES exam_records(id) ON DELETE CASCADE,
        FOREIGN KEY (question_id) REFERENCES questions(id) ON DELETE CASCADE)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS analysis_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER NOT NULL,
        record_date TEXT DEFAULT '', weakness_tag TEXT DEFAULT '',
        score INTEGER DEFAULT 0, raw_db_block TEXT DEFAULT '',
        feedback TEXT DEFAULT '', source TEXT DEFAULT 'manual',
        FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE)""")


def _migrate(cur):
    migrations = [
        ("passages", "content_hash", "TEXT DEFAULT ''"),
        ("passages", "updated_at", "TEXT DEFAULT ''"),
        ("questions", "choices", "TEXT DEFAULT '[]'"),
        ("questions", "choice_type", "TEXT DEFAULT 'num'"),
        ("questions", "answer", "TEXT DEFAULT ''"),
        ("questions", "explanation", "TEXT DEFAULT ''"),
        ("questions", "difficulty", "INTEGER DEFAULT 3"),
        ("students", "target", "TEXT DEFAULT ''"),
        ("students", "level", "TEXT DEFAULT ''"),
        ("students", "grade", "TEXT DEFAULT ''"),
        ("students", "created_at", "TEXT DEFAULT ''"),
        ("analysis_records", "source", "TEXT DEFAULT 'manual'"),
        ("analysis_records", "raw_db_block", "TEXT DEFAULT ''"),
    ]
    for table, col, ctype in migrations:
        try:
            cur.execute(f"PRAGMA table_info({table})")
            existing = {row[1] for row in cur.fetchall()}
            if col not in existing:
                cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ctype}")
        except sqlite3.OperationalError:
            pass


def db_stats(cfg):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        p = cur.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
        q = cur.execute("SELECT COUNT(*) FROM questions").fetchone()[0]
        s = cur.execute("SELECT COUNT(*) FROM students").fetchone()[0]
        e = cur.execute("SELECT COUNT(*) FROM exam_records").fetchone()[0]
        return {"passages": p, "questions": q, "students": s, "exams": e}
