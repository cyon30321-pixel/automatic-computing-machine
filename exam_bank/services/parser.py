"""
텍스트 파싱 엔진 — 붙여넣기 텍스트 → 지문/문항 구조화
"""

import re


def parse_exam_text(raw_text):
    blocks = re.split(r"\n(?=\s*\[\d+\]|\s*다음\s*글을\s*읽고)", "\n" + raw_text)
    parsed = []
    for block in blocks:
        block = block.strip()
        if len(block) < 15:
            continue
        parts = re.split(r"\n\s*\n(?=\d+\.\s)", block)
        passage = re.sub(
            r"^(?:\[\d+\]\s*)?(?:다음\s*글을\s*읽고\s*물음에\s*답하시오\.?)?",
            "", parts[0].strip(),
        ).strip()
        tag_m = re.match(r"^(\[\d+\])\s*", parts[0].strip())
        tag = tag_m.group(1) if tag_m else ""
        questions_raw = parts[1:] if len(parts) > 1 else []
        questions = []
        for qt in questions_raw:
            qt = qt.strip()
            if not qt:
                continue
            m = re.match(r"^(\d+)\.\s*(.*)", qt, re.DOTALL)
            questions.append({
                "q_num": m.group(1) if m else "-",
                "content": m.group(2) if m else qt,
                "choices": [], "answer": "", "explanation": "", "difficulty": 3,
            })
        parsed.append({"passage": passage, "questions": questions, "tag": tag})
    return parsed


def parse_question_with_choices(text):
    lines = text.strip().split("\n")
    question_lines = []
    choices = []
    choice_pattern = re.compile(r"^\s*[①②③④⑤]\s*")
    alpha_pattern = re.compile(r"^\s*[A-E][.)]\s*")
    for line in lines:
        if choice_pattern.match(line) or alpha_pattern.match(line):
            choices.append(line.strip())
        else:
            question_lines.append(line)
    return {"text": "\n".join(question_lines).strip(), "choices": choices}
