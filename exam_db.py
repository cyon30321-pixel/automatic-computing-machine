"""
exam_db.py  —  Question Bank Database Module
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DB 구조:
  db_root/
    items/<item_id>.json   ← passage 단위 1개 = 1파일
    backups/backup_*.zip
    index.json             ← 빠른 검색용 인덱스

의존성: Python 표준 라이브러리만 사용 (os, json, hashlib, datetime, shutil, zipfile, re)
"""
import os
import json
import hashlib
import datetime
import shutil
import zipfile
import re

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 상수
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCHEMA_VERSION = 1
CIRCLE_NUMS = "①②③④⑤"
ALPHA_CHOICES = ["A", "B", "C", "D", "E"]
VALID_ANSWERS_NUM = set(CIRCLE_NUMS)
VALID_ANSWERS_ALPHA = set(ALPHA_CHOICES)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 유틸: 정규화 / 검증
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def normalize_choices(choices):
    """선지를 항상 길이 5로 패딩."""
    result = list(choices or [])
    while len(result) < 5:
        result.append("")
    return result[:5]


def detect_c_type(choices):
    """선지 유형 감지: 'num'(①-⑤), 'alpha'(A-E), 기본값 'num'."""
    for ch in (choices or []):
        if ch and any(c in ch for c in CIRCLE_NUMS):
            return "num"
        if ch and any(ch.strip().upper().startswith(a) for a in ALPHA_CHOICES):
            return "alpha"
    return "num"


def validate_answer(answer, c_type):
    """정답 값 검증. 유효하면 반환, 아니면 None."""
    if not answer or not answer.strip():
        return None
    answer = answer.strip()
    if c_type == "alpha":
        return answer if answer.upper() in VALID_ANSWERS_ALPHA else None
    return answer if answer in VALID_ANSWERS_NUM else None


def compute_item_id(passage, questions):
    """passage + normalized questions + choices 기반 SHA1 해시 (16자)."""
    parts = [passage.strip()]
    for q in questions:
        parts.append(q.get("text", "").strip())
        for ch in q.get("choices", []):
            parts.append(str(ch).strip())
    raw = "|".join(parts)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 파싱 결과 ↔ DB 아이템 변환
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def parsed_to_db_items(parsed_questions, source_title="", tags=None):
    """parse_exam_text() 결과 → DB 아이템 리스트.
    같은 jesi(passage)를 공유하는 연속 문항은 하나의 아이템으로 그룹화."""
    if tags is None:
        tags = []
    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")

    groups = []
    current_passage = None
    current_questions = []

    for q in parsed_questions:
        jesi = q.get("jesi", "").strip()

        q_entry = {
            "num": q.get("num", ""),
            "text": q.get("text", ""),
            "choices": normalize_choices(q.get("choices", [])),
            "c_type": detect_c_type(q.get("choices", [])),
            "answer": None,
            "explanation": None,
        }
        # 재구성에 필요한 추가 필드 보존
        if q.get("bogi"):
            q_entry["bogi"] = q["bogi"]
        if q.get("tables"):
            q_entry["tables"] = q["tables"]
        if q.get("graph_tag"):
            q_entry["graph_tag"] = q["graph_tag"]

        if jesi:
            if jesi == current_passage:
                current_questions.append(q_entry)
            else:
                if current_questions:
                    groups.append({"passage": current_passage or "", "questions": current_questions})
                current_passage = jesi
                current_questions = [q_entry]
        else:
            if current_questions:
                groups.append({"passage": current_passage or "", "questions": current_questions})
            current_passage = ""
            current_questions = [q_entry]
            groups.append({"passage": "", "questions": current_questions})
            current_passage = None
            current_questions = []

    if current_questions:
        groups.append({"passage": current_passage or "", "questions": current_questions})

    items = []
    for grp in groups:
        item_id = compute_item_id(grp["passage"], grp["questions"])
        items.append({
            "schema_version": SCHEMA_VERSION,
            "item_id": item_id,
            "created_at": now,
            "updated_at": now,
            "source_title": source_title,
            "tags": list(tags),
            "passage": grp["passage"],
            "questions": grp["questions"],
        })
    return items


def db_items_to_exam_data(items):
    """DB 아이템 리스트 → create_exam_docx()가 요구하는 형식으로 변환.
    반환: [{"num", "text", "tables", "bogi", "jesi", "graph_tag", "choices"}, ...]"""
    result = []
    num_counter = 1
    for item in items:
        for q in item.get("questions", []):
            entry = {
                "num": str(num_counter),
                "text": q.get("text", ""),
                "tables": q.get("tables", []),
                "bogi": q.get("bogi", ""),
                "jesi": item.get("passage", ""),
                "graph_tag": q.get("graph_tag"),
                "choices": q.get("choices", []),
            }
            result.append(entry)
            num_counter += 1
    return result


def db_items_to_answer_text(items):
    """DB 아이템 리스트 → 정답/해설 텍스트 생성."""
    lines = []
    num_counter = 1
    for item in items:
        for q in item.get("questions", []):
            ans = q.get("answer") or "(미입력)"
            line = f"{num_counter}번 정답: {ans}"
            expl = q.get("explanation", "")
            if expl:
                line += f"\n[해설] {expl}"
            lines.append(line)
            num_counter += 1
    return "\n\n".join(lines)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# DB 디렉터리 관리
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def ensure_db_dirs(db_root):
    os.makedirs(os.path.join(db_root, "items"), exist_ok=True)
    os.makedirs(os.path.join(db_root, "backups"), exist_ok=True)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CRUD
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def save_item(db_root, item):
    """DB 아이템 저장. 동일 item_id 존재 시 merge (기존 정답/해설 보존)."""
    ensure_db_dirs(db_root)
    path = os.path.join(db_root, "items", f"{item['item_id']}.json")

    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                existing = json.load(f)
            for i, q in enumerate(item.get("questions", [])):
                if i < len(existing.get("questions", [])):
                    eq = existing["questions"][i]
                    if q.get("answer") is None and eq.get("answer"):
                        q["answer"] = eq["answer"]
                    if q.get("explanation") is None and eq.get("explanation"):
                        q["explanation"] = eq["explanation"]
            item["created_at"] = existing.get("created_at", item["created_at"])
            existing_tags = set(existing.get("tags", []))
            new_tags = set(item.get("tags", []))
            item["tags"] = sorted(existing_tags | new_tags)
        except (json.JSONDecodeError, OSError):
            pass

    item["updated_at"] = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(item, f, ensure_ascii=False, indent=2)
    return item["item_id"]


def load_item(db_root, item_id):
    path = os.path.join(db_root, "items", f"{item_id}.json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return None


def delete_item(db_root, item_id):
    path = os.path.join(db_root, "items", f"{item_id}.json")
    if os.path.exists(path):
        os.remove(path)
        return True
    return False


def list_all_items(db_root):
    items_dir = os.path.join(db_root, "items")
    if not os.path.isdir(items_dir):
        return []
    items = []
    for fname in sorted(os.listdir(items_dir)):
        if fname.endswith(".json"):
            try:
                with open(os.path.join(items_dir, fname), "r", encoding="utf-8") as f:
                    items.append(json.load(f))
            except (json.JSONDecodeError, OSError):
                continue
    return items


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 인덱스 관리
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def _build_search_text(item):
    parts = [
        item.get("source_title", ""),
        " ".join(item.get("tags", [])),
        item.get("passage", ""),
    ]
    for q in item.get("questions", []):
        parts.append(q.get("text", ""))
        parts.extend(q.get("choices", []))
    return " ".join(parts).lower()


def build_index(db_root):
    """모든 아이템으로부터 index.json 재생성."""
    items = list_all_items(db_root)
    index = {"version": 1, "items": {}}
    for item in items:
        iid = item.get("item_id", "")
        if not iid:
            continue
        index["items"][iid] = {
            "source_title": item.get("source_title", ""),
            "tags": item.get("tags", []),
            "updated_at": item.get("updated_at", ""),
            "question_count": len(item.get("questions", [])),
            "search_text": _build_search_text(item),
        }
    index_path = os.path.join(db_root, "index.json")
    ensure_db_dirs(db_root)
    with open(index_path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)
    return index


def load_index(db_root):
    index_path = os.path.join(db_root, "index.json")
    if not os.path.exists(index_path):
        return build_index(db_root)
    try:
        with open(index_path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return build_index(db_root)


def update_index_entry(db_root, item):
    """단일 아이템의 인덱스 엔트리 업데이트."""
    index = load_index(db_root)
    iid = item.get("item_id", "")
    index["items"][iid] = {
        "source_title": item.get("source_title", ""),
        "tags": item.get("tags", []),
        "updated_at": item.get("updated_at", ""),
        "question_count": len(item.get("questions", [])),
        "search_text": _build_search_text(item),
    }
    index_path = os.path.join(db_root, "index.json")
    with open(index_path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


def remove_index_entry(db_root, item_id):
    index = load_index(db_root)
    index["items"].pop(item_id, None)
    index_path = os.path.join(db_root, "index.json")
    with open(index_path, "w", encoding="utf-8") as f:
        json.dump(index, f, ensure_ascii=False, indent=2)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 검색
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def search_items(db_root, query):
    """인덱스 기반 OR 검색: 공백으로 분리한 토큰 중 하나라도 매치하면 반환."""
    index = load_index(db_root)
    if not query or not query.strip():
        return list(index.get("items", {}).keys())
    tokens = query.lower().split()
    results = []
    for iid, entry in index.get("items", {}).items():
        search_text = entry.get("search_text", "")
        if any(token in search_text for token in tokens):
            results.append(iid)
    return results


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 백업 / 복원
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def create_backup(db_root, config_path=None):
    """db_root/backups/에 zip 백업 생성."""
    ensure_db_dirs(db_root)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(db_root, "backups", f"backup_{ts}.zip")

    with zipfile.ZipFile(backup_path, "w", zipfile.ZIP_DEFLATED) as zf:
        items_dir = os.path.join(db_root, "items")
        if os.path.isdir(items_dir):
            for fname in os.listdir(items_dir):
                if fname.endswith(".json"):
                    fpath = os.path.join(items_dir, fname)
                    zf.write(fpath, f"items/{fname}")
        index_path = os.path.join(db_root, "index.json")
        if os.path.exists(index_path):
            zf.write(index_path, "index.json")
        if config_path and os.path.exists(config_path):
            zf.write(config_path, "exam_maker_config.json")
    return backup_path


def restore_backup(db_root, zip_path):
    """백업 zip에서 복원. 기존 items를 안전 복사 후 덮어쓰기."""
    if not os.path.exists(zip_path):
        raise FileNotFoundError(f"백업 파일 없음: {zip_path}")
    ensure_db_dirs(db_root)

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safety_dir = os.path.join(db_root, f"items_before_restore_{ts}")
    items_dir = os.path.join(db_root, "items")
    if os.path.isdir(items_dir) and os.listdir(items_dir):
        shutil.copytree(items_dir, safety_dir)

    with zipfile.ZipFile(zip_path, "r") as zf:
        for member in zf.namelist():
            if member.startswith("items/") and member.endswith(".json"):
                target = os.path.join(db_root, member)
                os.makedirs(os.path.dirname(target), exist_ok=True)
                with zf.open(member) as src, open(target, "wb") as dst:
                    dst.write(src.read())
            elif member == "index.json":
                target = os.path.join(db_root, "index.json")
                with zf.open(member) as src, open(target, "wb") as dst:
                    dst.write(src.read())
    return safety_dir


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# JSON Export
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
def export_exam_json(parsed_questions, target_dir, filename, display_title=""):
    """파싱된 시험 데이터를 JSON으로 export."""
    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    passages = []
    current_passage = ""
    current_questions = []

    for q in parsed_questions:
        jesi = q.get("jesi", "").strip()
        q_entry = {
            "num": q.get("num", ""),
            "text": q.get("text", ""),
            "choices": normalize_choices(q.get("choices", [])),
            "c_type": detect_c_type(q.get("choices", [])),
            "answer": q.get("answer", None),
        }
        if jesi != current_passage:
            if current_questions:
                passages.append({"passage": current_passage, "questions": current_questions})
            current_passage = jesi
            current_questions = [q_entry]
        else:
            current_questions.append(q_entry)
    if current_questions:
        passages.append({"passage": current_passage, "questions": current_questions})

    export_data = {
        "exported_at": now,
        "display_title": display_title,
        "passages": passages,
    }
    os.makedirs(target_dir, exist_ok=True)
    export_path = os.path.join(target_dir, f"{filename}_exam.json")
    with open(export_path, "w", encoding="utf-8") as f:
        json.dump(export_data, f, ensure_ascii=False, indent=2)
    return export_path


def export_from_db_items(items, target_dir, filename, display_title=""):
    """선택한 DB 아이템들을 exam JSON으로 export."""
    now = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    passages = []
    for item in items:
        passages.append({
            "passage": item.get("passage", ""),
            "questions": [
                {
                    "num": q.get("num", ""),
                    "text": q.get("text", ""),
                    "choices": q.get("choices", []),
                    "c_type": q.get("c_type", "num"),
                    "answer": q.get("answer"),
                }
                for q in item.get("questions", [])
            ],
        })
    export_data = {
        "exported_at": now,
        "display_title": display_title,
        "passages": passages,
    }
    os.makedirs(target_dir, exist_ok=True)
    export_path = os.path.join(target_dir, f"{filename}_exam.json")
    with open(export_path, "w", encoding="utf-8") as f:
        json.dump(export_data, f, ensure_ascii=False, indent=2)
    return export_path
