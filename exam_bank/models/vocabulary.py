"""
단어장 모델 v7.0.1 — CRUD, 엑셀 import, 학습 통계, Soft Delete, 예문(Sentence) 관리
v7.0.1: bulk_add_words_with_days에서 example 필드 지원 (문자열→vocab_sentences 자동 등록)
"""

import datetime
import json
from exam_bank.models.database import db_conn


# ── 단어장(Book) CRUD ──

def add_vocab_book(cfg, name, category="기본단어장"):
    now = datetime.datetime.now().isoformat()
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO vocab_books (name, category, created_at) VALUES (?,?,?)",
            (name, category, now))
        return cur.lastrowid


def update_vocab_book(cfg, book_id, name=None, category=None):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        sets, params = [], []
        if name is not None:
            sets.append("name=?"); params.append(name)
        if category is not None:
            sets.append("category=?"); params.append(category)
        if sets:
            params.append(book_id)
            cur.execute(f"UPDATE vocab_books SET {','.join(sets)} WHERE id=?", params)


def delete_vocab_book(cfg, book_id, hard=False):
    """Soft Delete (기본) — is_active=0. hard=True면 물리적 삭제."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        if hard:
            cur.execute("DELETE FROM vocab_words WHERE unit_id IN (SELECT id FROM vocab_units WHERE book_id=?)", (book_id,))
            cur.execute("DELETE FROM vocab_units WHERE book_id=?", (book_id,))
            cur.execute("DELETE FROM vocab_books WHERE id=?", (book_id,))
        else:
            cur.execute("UPDATE vocab_books SET is_active=0 WHERE id=?", (book_id,))
            cur.execute("UPDATE vocab_units SET is_active=0 WHERE book_id=?", (book_id,))


def list_vocab_books(cfg, category=None, include_inactive=False):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        query = "SELECT * FROM vocab_books"
        params = []
        conditions = []
        if not include_inactive:
            conditions.append("is_active=1")
        if category:
            conditions.append("category=?")
            params.append(category)
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        query += " ORDER BY category, name"
        rows = cur.execute(query, params).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_vocab_book(cfg, book_id):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        row = cur.execute("SELECT * FROM vocab_books WHERE id=?", (book_id,)).fetchone()
        if not row:
            return None
        return dict(zip([d[0] for d in cur.description], row))


# ── 단원(Unit) CRUD ──

def add_vocab_unit(cfg, book_id, unit_name, sort_order=0):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO vocab_units (book_id, unit_name, sort_order) VALUES (?,?,?)",
            (book_id, unit_name, sort_order))
        unit_id = cur.lastrowid
        # 단어장 total_units 갱신
        cnt = cur.execute(
            "SELECT COUNT(*) FROM vocab_units WHERE book_id=? AND is_active=1",
            (book_id,)).fetchone()[0]
        cur.execute("UPDATE vocab_books SET total_units=? WHERE id=?", (cnt, book_id))
        return unit_id


def update_vocab_unit(cfg, unit_id, unit_name=None, sort_order=None):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        sets, params = [], []
        if unit_name is not None:
            sets.append("unit_name=?"); params.append(unit_name)
        if sort_order is not None:
            sets.append("sort_order=?"); params.append(sort_order)
        if sets:
            params.append(unit_id)
            cur.execute(f"UPDATE vocab_units SET {','.join(sets)} WHERE id=?", params)


def delete_vocab_unit(cfg, unit_id, hard=False):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        if hard:
            cur.execute("DELETE FROM vocab_words WHERE unit_id=?", (unit_id,))
            cur.execute("DELETE FROM vocab_units WHERE id=?", (unit_id,))
        else:
            cur.execute("UPDATE vocab_units SET is_active=0 WHERE id=?", (unit_id,))
        # book total_units 갱신
        row = cur.execute("SELECT book_id FROM vocab_units WHERE id=?", (unit_id,)).fetchone()
        if row:
            cnt = cur.execute(
                "SELECT COUNT(*) FROM vocab_units WHERE book_id=? AND is_active=1",
                (row[0],)).fetchone()[0]
            cur.execute("UPDATE vocab_books SET total_units=? WHERE id=?", (cnt, row[0]))


def list_vocab_units(cfg, book_id, include_inactive=False):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        query = "SELECT * FROM vocab_units WHERE book_id=?"
        params = [book_id]
        if not include_inactive:
            query += " AND is_active=1"
        query += " ORDER BY sort_order, id"
        rows = cur.execute(query, params).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_vocab_unit(cfg, unit_id):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        row = cur.execute("SELECT * FROM vocab_units WHERE id=?", (unit_id,)).fetchone()
        if not row:
            return None
        return dict(zip([d[0] for d in cur.description], row))


# ── 단어(Word) CRUD ──

def add_vocab_word(cfg, unit_id, english, korean, pos="", example="", sort_order=0):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO vocab_words (unit_id, english, korean, part_of_speech, example_sentence, sort_order) VALUES (?,?,?,?,?,?)",
            (unit_id, english, korean, pos, example, sort_order))
        word_id = cur.lastrowid
        _update_unit_word_count(cur, unit_id)
        return word_id


def update_vocab_word(cfg, word_id, english=None, korean=None, pos=None, example=None):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        sets, params = [], []
        if english is not None:
            sets.append("english=?"); params.append(english)
        if korean is not None:
            sets.append("korean=?"); params.append(korean)
        if pos is not None:
            sets.append("part_of_speech=?"); params.append(pos)
        if example is not None:
            sets.append("example_sentence=?"); params.append(example)
        if sets:
            params.append(word_id)
            cur.execute(f"UPDATE vocab_words SET {','.join(sets)} WHERE id=?", params)


def delete_vocab_word(cfg, word_id):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        row = cur.execute("SELECT unit_id FROM vocab_words WHERE id=?", (word_id,)).fetchone()
        # v7: 예문도 삭제 (CASCADE 미지원 환경 대비)
        cur.execute("DELETE FROM vocab_sentences WHERE word_id=?", (word_id,))
        cur.execute("DELETE FROM vocab_words WHERE id=?", (word_id,))
        if row:
            _update_unit_word_count(cur, row[0])


def list_vocab_words(cfg, unit_id):
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        rows = cur.execute(
            "SELECT * FROM vocab_words WHERE unit_id=? ORDER BY sort_order, id",
            (unit_id,)).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_words_by_unit_ids(cfg, unit_ids):
    """여러 단원의 단어를 한꺼번에 가져오기 (출제범위 조합용)"""
    if not unit_ids:
        return []
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        placeholders = ",".join("?" * len(unit_ids))
        rows = cur.execute(
            f"""SELECT w.*, u.unit_name, u.book_id, b.name as book_name
                FROM vocab_words w
                JOIN vocab_units u ON w.unit_id = u.id
                JOIN vocab_books b ON u.book_id = b.id
                WHERE w.unit_id IN ({placeholders})
                ORDER BY u.sort_order, u.id, w.sort_order, w.id""",
            unit_ids).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def _update_unit_word_count(cur, unit_id):
    cnt = cur.execute(
        "SELECT COUNT(*) FROM vocab_words WHERE unit_id=?", (unit_id,)).fetchone()[0]
    cur.execute("UPDATE vocab_units SET word_count=? WHERE id=?", (cnt, unit_id))


# ── v7.0 예문(Sentence) CRUD ──

def add_sentence(cfg, word_id, sentence_en, sentence_ko="", target_form="",
                 difficulty=1, source="직접입력"):
    """단어에 예문 1개 추가. target_form 미지정 시 단어 원형 사용."""
    now = datetime.datetime.now().isoformat()
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        if not target_form:
            row = cur.execute("SELECT english FROM vocab_words WHERE id=?", (word_id,)).fetchone()
            target_form = row[0] if row else ""
        cur.execute(
            """INSERT INTO vocab_sentences
               (word_id, sentence_en, sentence_ko, target_form, difficulty, source, created_at)
               VALUES (?,?,?,?,?,?,?)""",
            (word_id, sentence_en.strip(), sentence_ko.strip(),
             target_form.strip(), difficulty, source, now))
        return cur.lastrowid


def update_sentence(cfg, sentence_id, sentence_en=None, sentence_ko=None,
                    target_form=None, difficulty=None, source=None):
    """예문 수정"""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        sets, params = [], []
        if sentence_en is not None:
            sets.append("sentence_en=?"); params.append(sentence_en.strip())
        if sentence_ko is not None:
            sets.append("sentence_ko=?"); params.append(sentence_ko.strip())
        if target_form is not None:
            sets.append("target_form=?"); params.append(target_form.strip())
        if difficulty is not None:
            sets.append("difficulty=?"); params.append(difficulty)
        if source is not None:
            sets.append("source=?"); params.append(source)
        if sets:
            params.append(sentence_id)
            cur.execute(f"UPDATE vocab_sentences SET {','.join(sets)} WHERE id=?", params)


def delete_sentence(cfg, sentence_id):
    """예문 1개 삭제"""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM vocab_sentences WHERE id=?", (sentence_id,))


def list_sentences(cfg, word_id, difficulty=None):
    """단어의 예문 목록. difficulty 지정 시 해당 난이도만."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        query = "SELECT * FROM vocab_sentences WHERE word_id=?"
        params = [word_id]
        if difficulty is not None:
            query += " AND difficulty=?"
            params.append(difficulty)
        query += " ORDER BY difficulty, id"
        rows = cur.execute(query, params).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]


def get_sentences_by_word_ids(cfg, word_ids, difficulty=None):
    """여러 단어의 예문을 한꺼번에 가져오기 (시험지 생성용)"""
    if not word_ids:
        return {}
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        placeholders = ",".join("?" * len(word_ids))
        query = f"SELECT * FROM vocab_sentences WHERE word_id IN ({placeholders})"
        params = list(word_ids)
        if difficulty is not None:
            query += " AND difficulty=?"
            params.append(difficulty)
        query += " ORDER BY word_id, difficulty, id"
        rows = cur.execute(query, params).fetchall()
        cols = [d[0] for d in cur.description]
        result = {}
        for r in rows:
            d = dict(zip(cols, r))
            wid = d["word_id"]
            if wid not in result:
                result[wid] = []
            result[wid].append(d)
        return result


def bulk_add_sentences(cfg, word_id, sentences_list, source="AI생성"):
    """단어 1개에 예문 여러 개 일괄 추가.
    sentences_list: [{"en": "...", "ko": "...", "target": "...", "diff": 1}, ...]
    """
    now = datetime.datetime.now().isoformat()
    count = 0
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        # target_form 기본값용 원형
        row = cur.execute("SELECT english FROM vocab_words WHERE id=?", (word_id,)).fetchone()
        default_form = row[0] if row else ""
        for s in sentences_list:
            en = s.get("en", "").strip()
            if not en:
                continue
            ko = s.get("ko", "").strip()
            target = s.get("target", default_form).strip()
            diff = s.get("diff", s.get("difficulty", 1))
            try:
                diff = int(diff)
            except (ValueError, TypeError):
                diff = 1
            diff = max(1, min(3, diff))  # 1~3 범위 제한
            cur.execute(
                """INSERT INTO vocab_sentences
                   (word_id, sentence_en, sentence_ko, target_form, difficulty, source, created_at)
                   VALUES (?,?,?,?,?,?,?)""",
                (word_id, en, ko, target, diff, source, now))
            count += 1
    return count


# ── 일괄 단어 등록 (JSON/엑셀) ──

def bulk_add_words(cfg, unit_id, words_list):
    """
    words_list: [{"english": "...", "korean": "...", "pos": "..."}, ...]
    반환: 등록 건수
    """
    count = 0
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        for i, w in enumerate(words_list):
            eng = w.get("english", "").strip()
            kor = w.get("korean", "").strip()
            if not eng:
                continue
            pos = w.get("pos", w.get("part_of_speech", "")).strip()
            example = w.get("example", w.get("example_sentence", "")).strip()
            cur.execute(
                "INSERT INTO vocab_words (unit_id, english, korean, part_of_speech, example_sentence, sort_order) VALUES (?,?,?,?,?,?)",
                (unit_id, eng, kor, pos, example, i + 1))
            count += 1
        _update_unit_word_count(cur, unit_id)
    return count


def bulk_add_words_with_days(cfg, book_id, day_words_dict):
    """
    여러 Day 한꺼번에 등록.
    day_words_dict: {"Day 01": [{"english": ..., "korean": ..., "example": "...", "sentences": [...]}, ...], ...}
    v7.0.1: example(문자열) 필드 지원 — vocab_words.example_sentence에 저장 + vocab_sentences에도 자동 등록
    v7.0: sentences(리스트) 필드가 있으면 예문도 자동 매칭 저장
    반환: {"Day 01": 40, "Day 02": 35, ...}  (Day별 등록 건수)
    """
    now = datetime.datetime.now().isoformat()
    results = {}
    total_sentences = 0
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        for sort_idx, (day_name, words) in enumerate(sorted(day_words_dict.items()), 1):
            # 기존 동일 이름 unit이 있으면 스킵
            existing = cur.execute(
                "SELECT id FROM vocab_units WHERE book_id=? AND unit_name=? AND is_active=1",
                (book_id, day_name)).fetchone()
            if existing:
                unit_id = existing[0]
            else:
                cur.execute(
                    "INSERT INTO vocab_units (book_id, unit_name, sort_order) VALUES (?,?,?)",
                    (book_id, day_name, sort_idx))
                unit_id = cur.lastrowid

            count = 0
            for i, w in enumerate(words):
                eng = w.get("english", "").strip()
                kor = w.get("korean", "").strip()
                if not eng:
                    continue
                pos = w.get("pos", w.get("part_of_speech", "")).strip()

                # ★ v7.0.1 수정: example 필드를 example_sentence 컬럼에도 저장
                example = w.get("example", w.get("example_sentence", "")).strip()

                cur.execute(
                    "INSERT INTO vocab_words (unit_id, english, korean, part_of_speech, example_sentence, sort_order) VALUES (?,?,?,?,?,?)",
                    (unit_id, eng, kor, pos, example, i + 1))
                word_id = cur.lastrowid
                count += 1

                # ★ v7.0.1 수정: example 문자열이 있으면 vocab_sentences에도 자동 등록
                if example:
                    cur.execute(
                        """INSERT INTO vocab_sentences
                           (word_id, sentence_en, sentence_ko, target_form, difficulty, source, created_at)
                           VALUES (?,?,?,?,?,?,?)""",
                        (word_id, example, "", eng, 1, "JSON가져오기", now))
                    total_sentences += 1

                # v7.0: sentences 리스트 필드가 있으면 추가 예문 저장
                sentences = w.get("sentences", [])
                if sentences and isinstance(sentences, list):
                    for s in sentences:
                        s_en = s.get("en", "").strip()
                        if not s_en:
                            continue
                        s_ko = s.get("ko", "").strip()
                        target = s.get("target", eng).strip()
                        diff = s.get("diff", s.get("difficulty", 1))
                        try:
                            diff = int(diff)
                        except (ValueError, TypeError):
                            diff = 1
                        diff = max(1, min(3, diff))
                        cur.execute(
                            """INSERT INTO vocab_sentences
                               (word_id, sentence_en, sentence_ko, target_form, difficulty, source, created_at)
                               VALUES (?,?,?,?,?,?,?)""",
                            (word_id, s_en, s_ko, target, diff, "AI생성", now))
                        total_sentences += 1

            _update_unit_word_count(cur, unit_id)
            results[day_name] = count

        # book total_units 갱신
        cnt = cur.execute(
            "SELECT COUNT(*) FROM vocab_units WHERE book_id=? AND is_active=1",
            (book_id,)).fetchone()[0]
        cur.execute("UPDATE vocab_books SET total_units=? WHERE id=?", (cnt, book_id))

    # 예문 총 개수를 결과에 포함 (UI 표시용)
    if total_sentences > 0:
        results["__sentences_total__"] = total_sentences
    return results


def import_words_from_excel(cfg, book_id, filepath, day_column="day", eng_column="english", kor_column="korean", pos_column="pos"):
    """
    엑셀/CSV에서 단어 일괄 등록.
    컬럼명은 유연하게 매칭.
    """
    import os
    ext = os.path.splitext(filepath)[1].lower()

    rows = []
    if ext in (".xlsx", ".xls"):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(filepath, read_only=True)
            ws = wb.active
            headers = [str(c.value or "").strip().lower() for c in next(ws.iter_rows(min_row=1, max_row=1))]
            for row in ws.iter_rows(min_row=2, values_only=True):
                rows.append(dict(zip(headers, [str(v or "").strip() for v in row])))
            wb.close()
        except ImportError:
            raise ImportError("openpyxl 패키지가 필요합니다. pip install openpyxl")
    elif ext == ".csv":
        import csv
        with open(filepath, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for row in reader:
                rows.append({k.strip().lower(): v.strip() for k, v in row.items()})
    elif ext == ".tsv":
        import csv
        with open(filepath, "r", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter="\t")
            for row in reader:
                rows.append({k.strip().lower(): v.strip() for k, v in row.items()})
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext}")

    # 컬럼 매핑 — ★ v7.0.1: example 컬럼도 전달
    day_words = {}
    for row in rows:
        day = row.get(day_column, row.get("day", row.get("unit", "Day 01")))
        eng = row.get(eng_column, row.get("english", row.get("word", "")))
        kor = row.get(kor_column, row.get("korean", row.get("meaning", "")))
        pos = row.get(pos_column, row.get("pos", row.get("part_of_speech", "")))
        example = row.get("example", row.get("example_sentence", ""))
        if not eng:
            continue
        if day not in day_words:
            day_words[day] = []
        day_words[day].append({"english": eng, "korean": kor, "pos": pos, "example": example})

    return bulk_add_words_with_days(cfg, book_id, day_words)


# ── 단어 시험 기록 ──

def create_vocab_exam(cfg, student_id, unit_ids, word_ids, exam_type="eng_to_kor", exam_title=""):
    """시험 기록 생성. word_ids: 출제된 단어 ID 리스트 (순서대로)"""
    now = datetime.datetime.now()
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        cur.execute(
            """INSERT INTO vocab_exam_records
               (student_id, exam_date, exam_title, total_words, exam_type, unit_ids_json, created_at)
               VALUES (?,?,?,?,?,?,?)""",
            (student_id, now.strftime("%Y-%m-%d"), exam_title,
             len(word_ids), exam_type, json.dumps(unit_ids), now.isoformat()))
        exam_id = cur.lastrowid
        for idx, wid in enumerate(word_ids, 1):
            cur.execute(
                "INSERT INTO vocab_exam_items (vocab_exam_id, word_id, display_order) VALUES (?,?,?)",
                (exam_id, wid, idx))
        return exam_id


def submit_vocab_answers(cfg, vocab_exam_id, answers):
    """
    answers: {word_id: student_answer_text, ...}
    정답 비교 후 is_correct 업데이트
    """
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        exam = cur.execute(
            "SELECT exam_type FROM vocab_exam_records WHERE id=?",
            (vocab_exam_id,)).fetchone()
        if not exam:
            return None
        exam_type = exam[0]

        items = cur.execute(
            """SELECT vei.id, vei.word_id, w.english, w.korean
               FROM vocab_exam_items vei
               JOIN vocab_words w ON vei.word_id = w.id
               WHERE vei.vocab_exam_id=?""",
            (vocab_exam_id,)).fetchall()

        total, correct = 0, 0
        for item_id, word_id, eng, kor in items:
            student_ans = answers.get(word_id, answers.get(str(word_id), "")).strip()
            if not student_ans:
                continue
            total += 1
            # 정답 비교
            if exam_type == "eng_to_kor":
                correct_ans = kor
            else:
                correct_ans = eng
            is_correct = 1 if student_ans.lower() == correct_ans.lower().strip() else 0
            if is_correct:
                correct += 1
            cur.execute(
                "UPDATE vocab_exam_items SET student_answer=?, is_correct=? WHERE id=?",
                (student_ans, is_correct, item_id))

        cur.execute(
            "UPDATE vocab_exam_records SET correct_count=? WHERE id=?",
            (correct, vocab_exam_id))
    return {"total": total, "correct": correct, "wrong": total - correct}


# ── 통계 ──

def get_student_vocab_stats(cfg, student_id, unit_id=None):
    """학생별 단어 시험 통계"""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        if unit_id:
            rows = cur.execute(
                """SELECT * FROM vocab_exam_records
                   WHERE student_id=? AND unit_ids_json LIKE ?
                   ORDER BY exam_date DESC""",
                (student_id, f"%{unit_id}%")).fetchall()
        else:
            rows = cur.execute(
                "SELECT * FROM vocab_exam_records WHERE student_id=? ORDER BY exam_date DESC",
                (student_id,)).fetchall()
        if not rows:
            return {"count": 0, "last_date": None, "best_score": 0, "avg_score": 0}
        cols = [d[0] for d in cur.description]
        records = [dict(zip(cols, r)) for r in rows]
        scores = []
        for r in records:
            if r["total_words"] > 0:
                scores.append(round(r["correct_count"] / r["total_words"] * 100, 1))
        return {
            "count": len(records),
            "last_date": records[0]["exam_date"] if records else None,
            "best_score": max(scores) if scores else 0,
            "avg_score": round(sum(scores) / len(scores), 1) if scores else 0,
        }


def get_vocab_unit_stats(cfg, book_id, student_id=None):
    """단원별 통계 (단어수, 시험횟수, 최근일시)"""
    units = list_vocab_units(cfg, book_id)
    results = []
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        for u in units:
            stat = {"unit_id": u["id"], "unit_name": u["unit_name"],
                    "word_count": u["word_count"], "exam_count": 0,
                    "last_date": None, "best_score": 0}
            if student_id:
                exams = cur.execute(
                    """SELECT exam_date, total_words, correct_count
                       FROM vocab_exam_records
                       WHERE student_id=? AND unit_ids_json LIKE ?
                       ORDER BY exam_date DESC""",
                    (student_id, f"%{u['id']}%")).fetchall()
                if exams:
                    stat["exam_count"] = len(exams)
                    stat["last_date"] = exams[0][0]
                    scores = [round(e[2] / e[1] * 100, 1) for e in exams if e[1] > 0]
                    stat["best_score"] = max(scores) if scores else 0
            results.append(stat)
    return results


def get_wrong_words(cfg, student_id, unit_ids=None):
    """학생이 틀린 단어만 추출 (오답 노트 / 재시험용)"""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        query = """
            SELECT DISTINCT w.id, w.english, w.korean, w.part_of_speech,
                   u.unit_name, b.name as book_name,
                   COUNT(CASE WHEN vei.is_correct=0 THEN 1 END) as wrong_count,
                   COUNT(CASE WHEN vei.is_correct=1 THEN 1 END) as correct_count
            FROM vocab_exam_items vei
            JOIN vocab_words w ON vei.word_id = w.id
            JOIN vocab_units u ON w.unit_id = u.id
            JOIN vocab_books b ON u.book_id = b.id
            JOIN vocab_exam_records ver ON vei.vocab_exam_id = ver.id
            WHERE ver.student_id=?
        """
        params = [student_id]
        if unit_ids:
            placeholders = ",".join("?" * len(unit_ids))
            query += f" AND w.unit_id IN ({placeholders})"
            params.extend(unit_ids)
        query += " GROUP BY w.id HAVING wrong_count > 0 ORDER BY wrong_count DESC"
        rows = cur.execute(query, params).fetchall()
        cols = [d[0] for d in cur.description]
        return [dict(zip(cols, r)) for r in rows]
