-- 통합 문제은행 v4.0 스키마
-- ================================

-- 지문 테이블
CREATE TABLE IF NOT EXISTS passages (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    content       TEXT NOT NULL,
    category1     TEXT DEFAULT '',
    category2     TEXT DEFAULT '',
    school_year   TEXT DEFAULT '',
    exam_year     TEXT DEFAULT '',
    exam_month    TEXT DEFAULT '',
    publisher     TEXT DEFAULT '',
    extra_tags    TEXT DEFAULT '',
    answer_text   TEXT DEFAULT '',
    content_hash  TEXT DEFAULT '',
    created_at    TEXT DEFAULT '',
    updated_at    TEXT DEFAULT ''
);

-- 문항 테이블
CREATE TABLE IF NOT EXISTS questions (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    passage_id    INTEGER NOT NULL,
    q_num         TEXT DEFAULT '-',
    content       TEXT DEFAULT '',
    choices       TEXT DEFAULT '[]',
    choice_type   TEXT DEFAULT 'num',
    answer        TEXT DEFAULT '',
    explanation   TEXT DEFAULT '',
    difficulty    INTEGER DEFAULT 3,
    usage_count   INTEGER DEFAULT 0,
    FOREIGN KEY (passage_id) REFERENCES passages(id) ON DELETE CASCADE
);

-- 학생 테이블
CREATE TABLE IF NOT EXISTS students (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    name          TEXT NOT NULL,
    target        TEXT DEFAULT '',
    level         TEXT DEFAULT '',
    grade         TEXT DEFAULT '',
    created_at    TEXT DEFAULT ''
);

-- 시험 출제 기록
CREATE TABLE IF NOT EXISTS exam_records (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id    INTEGER NOT NULL,
    exam_date     TEXT NOT NULL,
    exam_title    TEXT DEFAULT '',
    total_questions INTEGER DEFAULT 0,
    is_graded     INTEGER DEFAULT 0,
    score         REAL DEFAULT 0,
    created_at    TEXT DEFAULT '',
    FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE
);

-- 시험 문항 상세
CREATE TABLE IF NOT EXISTS exam_items (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    exam_id       INTEGER NOT NULL,
    question_id   INTEGER NOT NULL,
    display_order INTEGER DEFAULT 0,
    student_answer TEXT DEFAULT '',
    is_correct    INTEGER DEFAULT -1,
    FOREIGN KEY (exam_id) REFERENCES exam_records(id) ON DELETE CASCADE,
    FOREIGN KEY (question_id) REFERENCES questions(id) ON DELETE CASCADE
);

-- AI 분석 기록
CREATE TABLE IF NOT EXISTS analysis_records (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id    INTEGER NOT NULL,
    record_date   TEXT DEFAULT '',
    weakness_tag  TEXT DEFAULT '',
    score         INTEGER DEFAULT 0,
    raw_db_block  TEXT DEFAULT '',
    feedback      TEXT DEFAULT '',
    source        TEXT DEFAULT 'manual',
    FOREIGN KEY (student_id) REFERENCES students(id) ON DELETE CASCADE
);

-- 인덱스
CREATE INDEX IF NOT EXISTS idx_passages_hash ON passages(content_hash);
CREATE INDEX IF NOT EXISTS idx_questions_passage ON questions(passage_id);
CREATE INDEX IF NOT EXISTS idx_exam_records_student ON exam_records(student_id);
CREATE INDEX IF NOT EXISTS idx_exam_items_exam ON exam_items(exam_id);
CREATE INDEX IF NOT EXISTS idx_exam_items_question ON exam_items(question_id);
CREATE INDEX IF NOT EXISTS idx_analysis_student ON analysis_records(student_id);
