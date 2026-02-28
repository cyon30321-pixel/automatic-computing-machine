"""
백업/복원 서비스 — SQLite 백업, ZIP, JSON export
"""

import os
import json
import shutil
import zipfile
import datetime

from exam_bank.config import get_db_path
from exam_bank.models.database import db_conn


def backup_db(cfg):
    """DB 파일 백업. 반환: 백업 파일 경로."""
    src = get_db_path(cfg)
    if not os.path.exists(src):
        raise FileNotFoundError("DB 파일이 없습니다.")

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = src.replace(".db", f"_backup_{ts}.db")
    shutil.copy2(src, dst)
    return dst


def backup_db_zip(cfg):
    """DB + 설정을 ZIP으로 백업. 반환: ZIP 경로."""
    src = get_db_path(cfg)
    if not os.path.exists(src):
        raise FileNotFoundError("DB 파일이 없습니다.")

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    db_dir = cfg.get("db_dir", os.path.dirname(src))
    zip_path = os.path.join(db_dir, f"backup_{ts}.zip")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(src, os.path.basename(src))
    return zip_path


def restore_db(cfg, backup_path):
    """백업에서 DB 복원."""
    dst = get_db_path(cfg)

    if backup_path.endswith(".zip"):
        with zipfile.ZipFile(backup_path, "r") as zf:
            db_files = [n for n in zf.namelist() if n.endswith(".db")]
            if not db_files:
                raise ValueError("ZIP에 DB 파일이 없습니다.")
            with zf.open(db_files[0]) as src, open(dst, "wb") as f:
                f.write(src.read())
    else:
        shutil.copy2(backup_path, dst)


def list_backups(cfg):
    """백업 파일 목록."""
    db_dir = cfg.get("db_dir", "")
    if not os.path.isdir(db_dir):
        return []
    files = []
    for f in sorted(os.listdir(db_dir), reverse=True):
        if ("backup" in f) and (f.endswith(".db") or f.endswith(".zip")):
            full = os.path.join(db_dir, f)
            size = os.path.getsize(full)
            files.append({"name": f, "path": full, "size": size})
    return files


def export_json(cfg, passage_ids=None):
    """지문/문항을 JSON으로 export. passage_ids=None이면 전체."""
    with db_conn(cfg) as conn:
        cur = conn.cursor()
        if passage_ids:
            ph = ",".join("?" * len(passage_ids))
            passages = cur.execute(f"SELECT * FROM passages WHERE id IN ({ph})", passage_ids).fetchall()
        else:
            passages = cur.execute("SELECT * FROM passages").fetchall()

        p_cols = [d[0] for d in cur.description]
        result = []

        for p_row in passages:
            p = dict(zip(p_cols, p_row))
            qs = cur.execute("SELECT * FROM questions WHERE passage_id=?", (p["id"],)).fetchall()
            q_cols = [d[0] for d in cur.description]
            p["questions"] = [dict(zip(q_cols, q)) for q in qs]
            result.append(p)

    return result


def import_json(cfg, json_data):
    """JSON 데이터에서 지문/문항 가져오기."""
    from exam_bank.models.passage import create_passage

    count = 0
    for item in json_data:
        questions = []
        for q in item.get("questions", []):
            questions.append({
                "q_num": q.get("q_num", q.get("num", "-")),
                "content": q.get("content", q.get("text", "")),
                "choices": q.get("choices", []),
                "answer": q.get("answer", ""),
                "explanation": q.get("explanation", ""),
                "difficulty": q.get("difficulty", 3),
            })

        create_passage(
            cfg,
            content=item.get("content", item.get("passage", "")),
            category1=item.get("category1", ""),
            category2=item.get("category2", ""),
            school_year=item.get("school_year", ""),
            exam_year=item.get("exam_year", ""),
            exam_month=item.get("exam_month", ""),
            publisher=item.get("publisher", ""),
            extra_tags=item.get("extra_tags", ""),
            answer_text=item.get("answer_text", ""),
            questions=questions,
        )
        count += 1
    return count
