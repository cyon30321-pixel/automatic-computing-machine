"""
자동 출제 + 랜덤 셔플 서비스
"""

import random
from exam_bank.models.passage import auto_select_questions, get_questions_for_passage, get_passage


def auto_select(cfg, count=15, category1="", category2="",
                school_year="", min_diff=1, max_diff=5):
    return auto_select_questions(
        cfg, count=count, category1=category1, category2=category2,
        school_year=school_year, min_diff=min_diff, max_diff=max_diff
    )


def shuffle_question_ids(question_ids, seed=None):
    ids = list(question_ids)
    if seed is not None:
        random.seed(seed)
    random.shuffle(ids)
    return ids


def generate_multi_sets(question_ids, num_sets=2):
    labels = ["A형", "B형", "C형", "D형", "E형"]
    sets = []
    for i in range(min(num_sets, 5)):
        shuffled = shuffle_question_ids(question_ids, seed=i * 42 + 7)
        sets.append((labels[i], shuffled))
    return sets


def build_exam_data_from_cart(cfg, cart):
    exam_data = []
    all_qids = []
    for pid, qids in cart.items():
        p = get_passage(cfg, pid)
        if not p:
            continue
        qs = get_questions_for_passage(cfg, pid)
        if qids:
            qs = [q for q in qs if q["id"] in qids]
        for q in qs:
            all_qids.append(q["id"])
        exam_data.append({
            "pid": pid,
            "passage": p["content"],
            "questions": [{"num": q["q_num"], "text": q["content"],
                           "q_num": q["q_num"], "choices": q.get("choices", "[]")}
                          for q in qs],
            "info": f"{p.get('school_year', '')} {p.get('exam_year', '')} {p.get('exam_month', '')}".strip(),
            "answer_text": p.get("answer_text", ""),
        })
    return exam_data, all_qids
