"""
DB 코어 v6.0 — 연결 관리, 마이그레이션, 컨텍스트 매니저
v6.0: 단어장 테이블 4종, vocab_exam_items (오답 추적), is_active (Soft Delete)
"""

import os
import sqlite3
from contextlib import contextmanager

from exam_bank.config import get_db_path  # v5: question_bank_v5.db


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
    # ── 기존 테이블 (v5) ──
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

    # ── v6.0 단어장 테이블 ──
    cur.execute("""CREATE TABLE IF NOT EXISTS vocab_books (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        category TEXT DEFAULT '기본단어장',
        total_units INTEGER DEFAULT 0,
        is_active INTEGER DEFAULT 1,
        created_at TEXT DEFAULT '')""")

    cur.execute("""CREATE TABLE IF NOT EXISTS vocab_units (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        book_id INTEGER NOT NULL,
        unit_name TEXT NOT NULL,
        sort_order INTEGER DEFAULT 0,
        word_count INTEGER DEFAULT 0,
        is_active INTEGER DEFAULT 1,
        FOREIGN KEY (book_id) REFERENCES vocab_books(id))""")

    cur.execute("""CREATE TABLE IF NOT EXISTS vocab_words (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        unit_id INTEGER NOT NULL,
        english TEXT NOT NULL,
        korean TEXT NOT NULL,
        part_of_speech TEXT DEFAULT '',
        example_sentence TEXT DEFAULT '',
        sort_order INTEGER DEFAULT 0,
        FOREIGN KEY (unit_id) REFERENCES vocab_units(id))""")

    cur.execute("""CREATE TABLE IF NOT EXISTS vocab_exam_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_id INTEGER NOT NULL,
        exam_date TEXT NOT NULL,
        exam_title TEXT DEFAULT '',
        total_words INTEGER DEFAULT 0,
        correct_count INTEGER DEFAULT 0,
        exam_type TEXT DEFAULT 'eng_to_kor',
        unit_ids_json TEXT DEFAULT '[]',
        created_at TEXT DEFAULT '',
        FOREIGN KEY (student_id) REFERENCES students(id))""")

    cur.execute("""CREATE TABLE IF NOT EXISTS vocab_exam_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        vocab_exam_id INTEGER NOT NULL,
        word_id INTEGER NOT NULL,
        student_answer TEXT DEFAULT '',
        is_correct INTEGER DEFAULT -1,
        display_order INTEGER DEFAULT 0,
        FOREIGN KEY (vocab_exam_id) REFERENCES vocab_exam_records(id),
        FOREIGN KEY (word_id) REFERENCES vocab_words(id))""")


def _migrate(cur):
    migrations = [
        # v5 migrations
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
        # v6 migrations — soft delete + vocab
        ("vocab_books", "is_active", "INTEGER DEFAULT 1"),
        ("vocab_units", "is_active", "INTEGER DEFAULT 1"),
        ("vocab_units", "word_count", "INTEGER DEFAULT 0"),
        ("vocab_exam_records", "unit_ids_json", "TEXT DEFAULT '[]'"),
        ("vocab_exam_records", "exam_title", "TEXT DEFAULT ''"),
        ("vocab_exam_records", "created_at", "TEXT DEFAULT ''"),
    ]
    for table, col, ctype in migrations:
        try:
            cur.execute(f"PRAGMA table_info({table})")
            existing = {row[1] for row in cur.fetchall()}
            if col not in existing:
                cur.execute(f"ALTER TABLE {table} ADD COLUMN {col} {ctype}")
        except sqlite3.OperationalError:
            pass

    # v6: 새 테이블이 없으면 생성
    _ensure_vocab_tables(cur)


def _ensure_vocab_tables(cur):
    """레거시 v5 DB에서 업그레이드 시 vocab 테이블이 없을 수 있으므로 보장"""
    tables = {r[0] for r in cur.execute(
        "SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
    if "vocab_books" not in tables:
        cur.execute("""CREATE TABLE vocab_books (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL, category TEXT DEFAULT '기본단어장',
            total_units INTEGER DEFAULT 0, is_active INTEGER DEFAULT 1,
            created_at TEXT DEFAULT '')""")
    if "vocab_units" not in tables:
        cur.execute("""CREATE TABLE vocab_units (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            book_id INTEGER NOT NULL, unit_name TEXT NOT NULL,
            sort_order INTEGER DEFAULT 0, word_count INTEGER DEFAULT 0,
            is_active INTEGER DEFAULT 1,
            FOREIGN KEY (book_id) REFERENCES vocab_books(id))""")
    if "vocab_words" not in tables:
        cur.execute("""CREATE TABLE vocab_words (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            unit_id INTEGER NOT NULL, english TEXT NOT NULL,
            korean TEXT NOT NULL, part_of_speech TEXT DEFAULT '',
            example_sentence TEXT DEFAULT '', sort_order INTEGER DEFAULT 0,
            FOREIGN KEY (unit_id) REFERENCES vocab_units(id))""")
    if "vocab_exam_records" not in tables:
        cur.execute("""CREATE TABLE vocab_exam_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL, exam_date TEXT NOT NULL,
            exam_title TEXT DEFAULT '', total_words INTEGER DEFAULT 0,
            correct_count INTEGER DEFAULT 0, exam_type TEXT DEFAULT 'eng_to_kor',
            unit_ids_json TEXT DEFAULT '[]', created_at TEXT DEFAULT '',
            FOREIGN KEY (student_id) REFERENCES students(id))""")
    if "vocab_exam_items" not in tables:
        cur.execute("""CREATE TABLE vocab_exam_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            vocab_exam_id INTEGER NOT NULL, word_id INTEGER NOT NULL,
            student_answer TEXT DEFAULT '', is_correct INTEGER DEFAULT -1,
            display_order INTEGER DEFAULT 0,
            FOREIGN KEY (vocab_exam_id) REFERENCES vocab_exam_records(id),
            FOREIGN KEY (word_id) REFERENCES vocab_words(id))""")


def db_stats(cfg):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        def _count(tbl):
            try:
                return cur.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
            except Exception:
                return 0
        return {
            "passages": _count("passages"),
            "questions": _count("questions"),
            "students": _count("students"),
            "exams": _count("exam_records"),
            "vocab_books": _count("vocab_books"),
            "vocab_words": _count("vocab_words"),
        }
